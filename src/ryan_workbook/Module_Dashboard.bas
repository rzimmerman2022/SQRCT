Option Explicit

' --- Constants ---
Private Const DASHBOARD_SHEET_NAME As String = "SQRCT Dashboard"
Private Const USEREDITS_SHEET_NAME As String = "UserEdits"
Private Const USEREDITSLOG_SHEET_NAME As String = "UserEditsLog"
Private Const TEXT_ONLY_SHEET_NAME As String = "SQRCT Dashboard (Text-Only)" ' For Text-Only copy
Private Const MASTER_QUOTES_FINAL_SOURCE As String = "MasterQuotes_Final" ' PQ query/table for A-L
Private Const PQ_LATEST_LOCATION_SHEET As String = "DocNum_LatestLocation" ' PQ output sheet for Workflow Location
Private Const PQ_LATEST_LOCATION_TABLE As String = "DocNum_LatestLocation" ' PQ output table for Workflow Location
Private Const PQ_DOCNUM_COL_NAME As String = "PrimaryDocNumber"         ' Source DocNum Column Header for Workflow
Private Const PQ_LOCATION_COL_NAME As String = "MostRecent_FolderLocation" ' Source Location Column Header for Workflow
Private Const UE_COL_DOCNUM As String = "A"
Private Const UE_COL_PHASE As String = "B"
Private Const UE_COL_LASTCONTACT As String = "C"
Private Const UE_COL_COMMENTS As String = "D"
Private Const UE_COL_SOURCE As String = "E"
Private Const UE_COL_TIMESTAMP As String = "F"
Private Const DB_COL_WORKFLOW_LOCATION As String = "J" ' Populated by PopulateWorkflowLocation from PQ_LATEST_LOCATION
Private Const DB_COL_MISSING_QUOTE As String = "K" ' Populated from MasterQuotes_Final[AutoNote]
Private Const DB_COL_PHASE As String = "L" ' Populated from MasterQuotes_Final[AutoStage] initially, then UserEdits
Private Const DB_COL_LASTCONTACT As String = "M" ' Populated from UserEdits
Private Const DB_COL_COMMENTS As String = "N" ' Populated from UserEdits
Public Const DEBUG_WorkflowLocation As Boolean = False ' <<< SET TO TRUE FOR DEBUGGING WORKFLOW ONLY


'===============================================================================
'                         0. CORE HELPER ROUTINES
'===============================================================================

'------------------------------------------------------------------------------
' CleanDocumentNumber - Standardizes DocNum for matching.
' Mirrors Power Query Step Q logic for consistency.
' *** Using robust version ***
'------------------------------------------------------------------------------
Public Function CleanDocumentNumber(ByVal raw As String) As String
    Dim cleaned As String
    Dim prefix As String, tail As String
    Dim num As Long, tailLen As Long

    ' 1. Trim and Uppercase
    cleaned = UCase$(Trim$(raw))

    ' 2. Handle empty or short strings
    If Len(cleaned) = 0 Then CleanDocumentNumber = "": Exit Function
    If Len(cleaned) <= 5 Then CleanDocumentNumber = cleaned: Exit Function ' Return as-is if 5 chars or less

    ' 3. Split into Prefix (5 chars) and Tail
    prefix = Left$(cleaned, 5)
    tail = Mid$(cleaned, 6)

    ' 4. Handle BSMOQ prefix exception (return as-is)
    If prefix = "BSMOQ" Then CleanDocumentNumber = cleaned: Exit Function

    ' 5. Pad numeric tail to 5 digits (or original length if longer)
    If VBA.isNumeric(tail) Then ' Use VBA.IsNumeric for clarity
        On Error Resume Next ' Handle potential overflow if tail is huge number
        num = Val(tail)
        If Err.Number <> 0 Then ' Val failed (e.g., too large)
            Err.Clear
            ' Keep original tail if Val fails
        Else
            On Error GoTo 0 ' Restore error handling
            tailLen = Len(tail)
            ' Ensure padding doesn't truncate if original tail was > 5 digits and numeric
            tail = Right$("00000" & CStr(num), IIf(tailLen < 5, 5, tailLen))
        End If
        On Error GoTo 0
    Else
        ' Keep original tail if it's not purely numeric
    End If

    ' 6. Combine and return
    CleanDocumentNumber = prefix & tail
End Function

'------------------------------------------------------------------------------
' LogUserEditsOperation – Timestamped entry to hidden log sheet
' *** Using robust version ***
'------------------------------------------------------------------------------
Public Sub LogUserEditsOperation(msg As String)
    Dim wsLog As Worksheet, r As Long
    On Error Resume Next ' Prevent error if sheet exists but is protected differently
    Set wsLog = ThisWorkbook.Sheets(USEREDITSLOG_SHEET_NAME)
    On Error GoTo 0 ' Restore error handling

    If wsLog Is Nothing Then
        On Error Resume Next ' Prevent error if workbook is protected
        Set wsLog = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        If wsLog Is Nothing Then Exit Sub ' Failed to add sheet
        On Error GoTo 0
        wsLog.Name = USEREDITSLOG_SHEET_NAME
        wsLog.Range("A1:C1").Value = Array("Timestamp", "Workbook", "Operation")
        wsLog.Range("A1:C1").Font.Bold = True
        wsLog.Visible = xlSheetHidden ' Hide it by default
    End If

    On Error Resume Next ' Avoid error if log sheet is protected
    r = wsLog.Cells(wsLog.rows.Count, "A").End(xlUp).Row + 1
    If r < 2 Then r = 2 ' Ensure we start writing at row 2 if sheet was empty
    wsLog.Cells(r, "A").Value = Format$(Now, "yyyy-mm-dd hh:nn:ss")
    wsLog.Cells(r, "B").Value = Module_Identity.GetWorkbookIdentity() ' Use Identity Module
    wsLog.Cells(r, "C").Value = msg
    If Err.Number <> 0 Then Debug.Print "Error writing to UserEditsLog: " & Err.Description: Err.Clear
    On Error GoTo 0
End Sub

'------------------------------------------------------------------------------
' LoadUserEditsToDictionary – Builds Doc# -> SheetRow dictionary from UserEdits
' Uses CleanDocumentNumber for keys.
' *** Using robust version ***
'------------------------------------------------------------------------------
Public Function LoadUserEditsToDictionary(wsEdits As Worksheet) As Object
    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")
    dict.CompareMode = vbTextCompare ' Use case-insensitive comparison for dictionary keys

    If wsEdits Is Nothing Then
        LogUserEditsOperation "LoadUserEditsDict: UserEdits sheet object is Nothing."
        Set LoadUserEditsToDictionary = dict ' Return empty dictionary
        Exit Function
    End If

    Dim lastRow As Long, r As Long
    Dim rawDocNum As String, cleanedDocNum As String
    Dim addedCount As Long, duplicateCount As Long, skippedCount As Long

    On Error Resume Next ' Handle error if sheet is protected or other access issues
    lastRow = wsEdits.Cells(wsEdits.rows.Count, UE_COL_DOCNUM).End(xlUp).Row
    If Err.Number <> 0 Then
        LogUserEditsOperation "LoadUserEditsDict: Error finding last row on '" & wsEdits.Name & "'. Error: " & Err.Description
        Set LoadUserEditsToDictionary = dict ' Return empty dictionary
        Exit Function
    End If
    On Error GoTo 0 ' Restore default error handling for this sub

    LogUserEditsOperation "LoadUserEditsDict: Checking UserEdits sheet '" & wsEdits.Name & "' rows 2 through " & lastRow & "."

    If lastRow > 1 Then ' Only process if there are potential data rows
        For r = 2 To lastRow ' Loop through actual data rows
            rawDocNum = CStr(wsEdits.Cells(r, UE_COL_DOCNUM).Value) ' Get raw value from UserEdits Col A
            cleanedDocNum = CleanDocumentNumber(rawDocNum)         ' Clean it using the helper function

            If cleanedDocNum <> "" Then ' Only process if cleaning results in a non-empty string
                If Not dict.Exists(cleanedDocNum) Then
                    dict.Add key:=cleanedDocNum, Item:=r ' Key = Cleaned DocNum, Item = Sheet Row Number
                    addedCount = addedCount + 1
                Else
                    duplicateCount = duplicateCount + 1
                    ' Log duplicates but keep the first one found
                    ' LogUserEditsOperation "LoadUserEditsDict: WARNING - Duplicate CLEANED DocNum '" & cleanedDocNum & "' found at UserEdits row " & r & ". Keeping mapping to first occurrence (Row " & dict(cleanedDocNum) & ")."
                End If
            Else
                skippedCount = skippedCount + 1
            End If
        Next r
        LogUserEditsOperation "LoadUserEditsDict: Processed " & lastRow - 1 & " rows. Added: " & addedCount & ", Duplicates Ignored: " & duplicateCount & ", Skipped (Blank): " & skippedCount & "."
    Else
       LogUserEditsOperation "LoadUserEditsDict: No data rows found on UserEdits sheet (lastRow=" & lastRow & ")."
    End If

    Set LoadUserEditsToDictionary = dict ' Return the populated (or empty) dictionary
End Function


'------------------------------------------------------------------------------
' IsMasterQuotesFinalPresent - Checks if the source exists (PQ, Table, or NR)
' *** Using robust version ***
'------------------------------------------------------------------------------
Public Function IsMasterQuotesFinalPresent() As Boolean
    Dim lo As ListObject
    Dim nm As Name
    Dim queryObj As Object
    Dim ws As Worksheet

    IsMasterQuotesFinalPresent = False
    On Error Resume Next ' Ignore errors during checks

    ' 1. Check for Power Query
    Set queryObj = Nothing: Err.Clear
    Set queryObj = ThisWorkbook.Queries(MASTER_QUOTES_FINAL_SOURCE)
    If Err.Number = 0 And Not queryObj Is Nothing Then
        IsMasterQuotesFinalPresent = True
        GoTo ExitPoint_IsPresent ' Found it
    End If

    ' 2. Check for Table (ListObject) on any sheet
    Set lo = Nothing: Err.Clear
    For Each ws In ThisWorkbook.Worksheets
        Set lo = ws.ListObjects(MASTER_QUOTES_FINAL_SOURCE)
        If Err.Number = 0 And Not lo Is Nothing Then
            IsMasterQuotesFinalPresent = True
            GoTo ExitPoint_IsPresent ' Found it
        End If
    Next ws

    ' 3. Check for Named Range
    Set nm = Nothing: Err.Clear
    Set nm = ThisWorkbook.Names(MASTER_QUOTES_FINAL_SOURCE)
    If Err.Number = 0 And Not nm Is Nothing Then
        ' Further check if the named range actually refers to a valid range
        Dim tempRange As Range
        Set tempRange = Nothing
        Set tempRange = nm.RefersToRange
        If Err.Number = 0 And Not tempRange Is Nothing Then
             IsMasterQuotesFinalPresent = True
             GoTo ExitPoint_IsPresent ' Found it
        End If
    End If

ExitPoint_IsPresent:
    On Error GoTo 0 ' Restore default error handling
    Set lo = Nothing: Set nm = Nothing: Set queryObj = Nothing: Set ws = Nothing: Set tempRange = Nothing
End Function


'===============================================================================
'              1. DASHBOARD CREATION / REFRESH MASTER ROUTINE
'===============================================================================

'------------------------------------------------------------------------------
' Button macros (Assign these to shapes/buttons on the dashboard)
' *** Kept from User's Provided Code ***
'------------------------------------------------------------------------------
Public Sub Button_RefreshDashboard_SaveAndRestoreEdits()
    ' Standard workflow: Saves dashboard edits (L-N) -> Refreshes A-L -> Restores all UserEdits (L-N)
    RefreshDashboard PreserveUserEdits:=False
End Sub

Public Sub Button_RefreshDashboard_PreserveUserEdits()
    ' Preserve workflow: Refreshes A-L -> Restores all UserEdits (L-N) -> Does NOT save current dashboard edits first
    RefreshDashboard PreserveUserEdits:=True
End Sub

'------------------------------------------------------------------------------
' RefreshDashboard — Master routine orchestrating the entire refresh
' *** Incorporating functional fixes into user's preferred structure ***
'------------------------------------------------------------------------------
Public Sub RefreshDashboard(Optional PreserveUserEdits As Boolean = False)

    Dim ws As Worksheet, wsEdits As Worksheet
    Dim lastRow As Long, lastRowEdits As Long
    Dim t_Start As Double, t_Populate As Double, t_Sort As Double, t_Freeze As Double
    Dim t_Workflow As Double, t_Restore As Double, t_Format As Double, t_TextOnly As Double ' Timers
    Dim userEditsDict As Object ' Dictionary for lookup
    Dim backupCreated As Boolean
    Dim calcState As XlCalculation: calcState = Application.Calculation ' Store current calculation state
    Dim eventsState As Boolean: eventsState = Application.EnableEvents ' Store current event state
    Dim currentSheet As Worksheet: Set currentSheet = ActiveSheet ' Remember active sheet

    t_Start = Timer ' Start overall timer
    LogUserEditsOperation "--------------------------------------------------"
    LogUserEditsOperation "Starting Dashboard Refresh. Mode: " & IIf(PreserveUserEdits, "PreserveUserEdits", "SaveAndRestore")

    ' --- Attempt pre-emptive backup ---
    backupCreated = CreateUserEditsBackup("RefreshDashboard_" & Format(Now, "yyyymmdd_hhmmss"))
    LogUserEditsOperation "Pre-refresh UserEdits backup created: " & backupCreated

    ' --- Error Handling & Application Settings ---
    On Error GoTo ErrorHandler ' Master error handler for the refresh process
    Application.ScreenUpdating = False
    Application.EnableEvents = False ' Turn off events during manipulation
    Application.Calculation = xlCalculationManual ' Manual calculation for speed
    Application.DisplayAlerts = False ' Suppress alerts like overwrite confirmations

    '--- STEP 1: Ensure UserEdits Sheet Exists and Get Reference ---
    SetupUserEditsSheet ' Creates or verifies the UserEdits sheet structure
    Set wsEdits = Nothing ' Ensure variable is clear before setting
    On Error Resume Next ' Handle potential error getting sheet reference
    Set wsEdits = ThisWorkbook.Sheets(USEREDITS_SHEET_NAME)
    On Error GoTo ErrorHandler ' Restore main error handler
    If wsEdits Is Nothing Then
        MsgBox "CRITICAL ERROR: Could not find or create the '" & USEREDITS_SHEET_NAME & "' sheet. Aborting refresh.", vbCritical, "Refresh Aborted"
        GoTo Cleanup ' Cannot proceed
    End If

    '--- STEP 2: Save Current Dashboard Edits (L-N) to UserEdits (if not preserving) ---
    If Not PreserveUserEdits Then
        LogUserEditsOperation "SaveAndRestore Mode: Saving current dashboard edits (L-N) to UserEdits sheet..."
        SaveUserEditsFromDashboard ' Saves L-N from Dashboard to UserEdits A-F
        LogUserEditsOperation "Finished saving dashboard edits."
    Else
        LogUserEditsOperation "PreserveUserEdits Mode: Skipping save of dashboard edits."
    End If

    '--- STEP 3: Get or Create Dashboard Sheet & Prepare Layout ---
    Set ws = GetOrCreateDashboardSheet(DASHBOARD_SHEET_NAME) ' Finds or creates the main dashboard sheet
    If ws Is Nothing Then
         MsgBox "CRITICAL ERROR: Could not find or create the '" & DASHBOARD_SHEET_NAME & "' sheet. Aborting refresh.", vbCritical, "Refresh Aborted"
         GoTo Cleanup ' Cannot proceed
    End If

    On Error Resume Next ' Ignore error if already unprotected
    ws.Unprotect
    On Error GoTo ErrorHandler ' Restore error handler

    CleanupDashboardLayout ws    ' Uses user's preferred version - Clears rows 4+, ensures Rows 1-3 exist
    InitializeDashboardLayout ws ' Uses user's preferred version - Sets headers A-N and initial widths

    '--- STEP 4: Populate Dashboard Columns A-L from MasterQuotes_Final ---
    LogUserEditsOperation "Populating dashboard columns A-L from '" & MASTER_QUOTES_FINAL_SOURCE & "'..."
    t_Populate = Timer
    If IsMasterQuotesFinalPresent Then
        ' *** Using robust PopulateMasterQuotesData sub ***
        PopulateMasterQuotesData ws ' Fills A-I (Core), K (AutoNote), L (AutoStage for Legacy) using formulas
    Else
        LogUserEditsOperation "WARNING: Source '" & MASTER_QUOTES_FINAL_SOURCE & "' not found. Dashboard columns A-L will be empty."
        MsgBox "Warning: Data source '" & MASTER_QUOTES_FINAL_SOURCE & "' not found." & vbCrLf & _
               "Dashboard columns A-L cannot be populated.", vbExclamation, "Data Source Missing"
    End If
    LogUserEditsOperation "Finished populating A-L. Time: " & Format(Timer - t_Populate, "0.00") & "s"

    ' Get last row *after* populating A-L
    lastRow = ws.Cells(ws.rows.Count, "A").End(xlUp).Row
    LogUserEditsOperation "Dashboard has " & IIf(lastRow < 4, 0, lastRow - 3) & " data rows after initial population (Row " & lastRow & ")."

    '--- STEP 5: Freeze Formulas (A-L) to Values *BEFORE* Sorting ---
    If lastRow >= 4 Then
        LogUserEditsOperation "Freezing formulas to values for columns A through L..."
        t_Freeze = Timer
        On Error Resume Next ' Handle potential errors during conversion
        With ws.Range("A4:" & DB_COL_PHASE & lastRow) ' Range A:L includes new Phase column
            .Value = .Value ' Convert formulas to static values
        End With
        If Err.Number <> 0 Then
             LogUserEditsOperation "Warning: Error during formula-to-value conversion (A:L): " & Err.Description
             Err.Clear ' Clear error and continue
        End If
        On Error GoTo ErrorHandler ' Restore main error handler
        LogUserEditsOperation "Finished freezing formulas. Time: " & Format(Timer - t_Freeze, "0.00") & "s"
    End If

    '--- STEP 6: Sort Dashboard Data ---
    If lastRow >= 5 Then ' Need at least 2 data rows to sort meaningfully
        LogUserEditsOperation "Sorting dashboard rows 4:" & lastRow & "..."
        t_Sort = Timer
        SortDashboardData ws, lastRow ' Uses user's preferred version
        LogUserEditsOperation "Finished sorting. Time: " & Format(Timer - t_Sort, "0.00") & "s"
    Else
        LogUserEditsOperation "Skipping sort (less than 2 data rows)."
    End If
    ' Recalculate lastRow *after* sorting
    lastRow = ws.Cells(ws.rows.Count, "A").End(xlUp).Row

    '--- STEP 7: Populate Workflow Location (Column J) ---
    If lastRow >= 4 Then
        LogUserEditsOperation "Populating Workflow Location (Column J)..."
        t_Workflow = Timer
        PopulateWorkflowLocation ws, lastRow ' Uses version with DEBUG toggle
        LogUserEditsOperation "Finished Workflow Location. Time: " & Format(Timer - t_Workflow, "0.00") & "s"
    Else
         LogUserEditsOperation "Skipping Workflow Location population (no data rows)."
    End If

        '--- STEP 8: Restore UserEdits (L-N) from UserEdits Sheet ---
        LogUserEditsOperation "Restoring UserEdits (Phase/Contact/Comments) to dashboard columns L-N..."
        t_Restore = Timer
        RestoreUserEditsToDashboard ws, wsEdits, lastRow
        LogUserEditsOperation "Finished UserEdits restoration. Time: " & Format(Timer - t_Restore, "0.00") & "s"

' *** REVISED STEP 8.5: AutoFit A:M THEN Set Fixed/Minimum Widths ***
        LogUserEditsOperation "Applying AutoFit to columns A:M then setting fixed/minimums..."
        Application.ScreenUpdating = False ' Ensure it's off
        On Error Resume Next ' Ignore errors during formatting

        ' 1. AutoFit Columns A through M only
        ws.Columns("A:" & DB_COL_LASTCONTACT).AutoFit ' AutoFit A-M (DB_COL_LASTCONTACT is M)

        ' 2. NOW, enforce MINIMUM widths AFTER AutoFit for button columns C, D, E
        If ws.Columns("C").ColumnWidth < 20 Then ws.Columns("C").ColumnWidth = 20 ' Min width for C
        If ws.Columns("D").ColumnWidth < 5 Then ws.Columns("D").ColumnWidth = 5   ' Min width for D
        If ws.Columns("E").ColumnWidth < 20 Then ws.Columns("E").ColumnWidth = 20 ' Min width for E

        ' 3. Set FIXED width for Comments column N AFTER AutoFit of other columns
        ws.Columns(DB_COL_COMMENTS).ColumnWidth = 45 ' Set Fixed width for Comments (N) - Adjust 45 if needed

        ' 4. Optional: Apply specific FIXED widths AFTER AutoFit for other columns if desired
        ' ws.Columns("A").ColumnWidth = 15 ' DocNum

        If Err.Number <> 0 Then LogUserEditsOperation "WARNING: Error during AutoFit/Width Setting.": Err.Clear
        On Error GoTo ErrorHandler ' Restore main handler
        Application.ScreenUpdating = True
        LogUserEditsOperation "Finished column AutoFit and width adjustments (N fixed)."
        ' *** END REVISED STEP 8.5 ***

        ' 1. AutoFit ALL columns first to adjust to content
        ws.Columns("A:" & DB_COL_COMMENTS).AutoFit ' AutoFit A-N

        ' 2. NOW, enforce MINIMUM widths for columns C, D, E AFTER AutoFit
        '    Adjust these minimums if needed based on your font/button size
        If ws.Columns("C").ColumnWidth < 20 Then ws.Columns("C").ColumnWidth = 20 ' Min width for C (Base for Button 1)
        If ws.Columns("D").ColumnWidth < 5 Then ws.Columns("D").ColumnWidth = 5   ' Min width for D (Spacer)
        If ws.Columns("E").ColumnWidth < 20 Then ws.Columns("E").ColumnWidth = 20 ' Min width for E (Base for Button 2)

        ' 3. Optional: Set specific FIXED widths AFTER AutoFit for other columns if desired
        ' ws.Columns("N").ColumnWidth = 40 ' Example: Ensure Comments is wide enough

        If Err.Number <> 0 Then LogUserEditsOperation "WARNING: Error during AutoFit/Width Setting for A:N.": Err.Clear
        On Error GoTo ErrorHandler ' Restore main handler
        Application.ScreenUpdating = True
        LogUserEditsOperation "Finished column AutoFit and minimum width adjustments."
        ' *** END REVISED STEP 8.5 ***

        '--- STEP 9: Apply Formatting and Protection ---
        ' (Rest of the code remains the same)
        LogUserEditsOperation "Applying formatting and protection..."
        t_Format = Timer
        If lastRow >= 4 Then
            ApplyColorFormatting ws, 4
            ApplyWorkflowLocationFormatting ws, 4
        End If
        ProtectUserColumns ws
        FreezeDashboard ws
        LogUserEditsOperation "Finished formatting/protection. Time: " & Format(Timer - t_Format, "0.00") & "s"

        '--- STEP 10: Final UI Updates (Timestamp, Buttons) ---
        SetupDashboardUI_EndRefresh ws

    '--- STEP 11: Create/Update Text-Only Copy ---
    LogUserEditsOperation "Creating/Updating Text-Only dashboard copy..."
    t_TextOnly = Timer
    CreateOrUpdateTextOnlySheet ws ' Use reviewed version
    LogUserEditsOperation "Finished Text-Only copy. Time: " & Format(Timer - t_TextOnly, "0.00") & "s"

    '--- STEP 12: Completion Message & Cleanup ---
    Dim msgText As String
    If PreserveUserEdits Then
        msgText = DASHBOARD_SHEET_NAME & " refreshed successfully!" & vbCrLf & vbCrLf & _
                  "User edits from the '" & USEREDITS_SHEET_NAME & "' sheet were applied." & vbCrLf & _
                  "(Dashboard edits were NOT saved back during this refresh.)"
    Else
        msgText = DASHBOARD_SHEET_NAME & " refreshed successfully!" & vbCrLf & vbCrLf & _
                  "Edits made on the dashboard were saved to '" & USEREDITS_SHEET_NAME & "'." & vbCrLf & _
                  "All edits were then restored to the dashboard."
    End If
    Application.DisplayAlerts = True ' Re-enable alerts before showing message box
    MsgBox msgText, vbInformation, "Dashboard Refresh Complete"
    Application.DisplayAlerts = False ' Turn back off before potential backup cleanup

    LogUserEditsOperation "Dashboard refresh completed successfully. Total time: " & Format(Timer - t_Start, "0.00") & "s"

    ' Clean up old backups if refresh was successful
    If backupCreated Then CleanupOldBackups ' Deletes backups older than specified days

Cleanup: ' Label for normal exit and error exit cleanup
    On Error Resume Next ' Prevent errors during cleanup itself
    ' --- Restore Application Settings ---
    Application.ScreenUpdating = True
    Application.Calculation = calcState ' Restore original calculation state
    Application.DisplayAlerts = True
    Application.EnableEvents = eventsState ' Restore original event state LAST
    ' --- Release Object Variables ---
    Set ws = Nothing
    Set wsEdits = Nothing
    Set userEditsDict = Nothing
    ' Activate the original sheet if it was changed
    If Not currentSheet Is Nothing Then
        If ActiveSheet.Name <> currentSheet.Name Then currentSheet.Activate
    End If
    Set currentSheet = Nothing

    LogUserEditsOperation "Cleanup complete."
    LogUserEditsOperation "=================================================="
    Exit Sub ' Normal exit

ErrorHandler: ' Master error handler
    Dim errNum As Long: errNum = Err.Number
    Dim errDesc As String: errDesc = Err.Description
    Dim errLine As Long: errLine = Erl ' Get line number where error occurred

    LogUserEditsOperation "!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!"
    LogUserEditsOperation "ERROR #" & errNum & " occurred in RefreshDashboard on or near line " & errLine & ":" & vbCrLf & errDesc
    LogUserEditsOperation "!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!"

    ' Attempt to restore from the backup created at the start
    If backupCreated Then
        LogUserEditsOperation "Attempting to restore UserEdits from pre-refresh backup due to error..."
        If RestoreUserEditsFromBackup() Then ' Tries to find most recent backup
            LogUserEditsOperation "UserEdits restore from backup SUCCEEDED."
            MsgBox "An error occurred during the refresh." & vbCrLf & vbCrLf & _
                   "Error: " & errDesc & vbCrLf & "(Error Code: " & errNum & ")" & vbCrLf & vbCrLf & _
                   "Your UserEdits sheet has been restored from the backup created before this refresh.", vbCritical, "Dashboard Refresh Error"
        Else
            LogUserEditsOperation "UserEdits restore from backup FAILED."
             MsgBox "An error occurred during the refresh." & vbCrLf & vbCrLf & _
                   "Error: " & errDesc & vbCrLf & "(Error Code: " & errNum & ")" & vbCrLf & vbCrLf & _
                   "ATTEMPT TO RESTORE USEREDITS FROM BACKUP FAILED. Please check manually for backup sheets ('" & USEREDITS_SHEET_NAME & "_Backup...').", vbCritical, "Dashboard Refresh Error"
        End If
    Else
         MsgBox "An error occurred during the refresh." & vbCrLf & vbCrLf & _
               "Error: " & errDesc & vbCrLf & "(Error Code: " & errNum & ")" & vbCrLf & vbCrLf & _
               "No pre-refresh backup was successfully created.", vbCritical, "Dashboard Refresh Error"
    End If

    ' Perform cleanup tasks even after error
    Resume Cleanup

End Sub


'================================================================================
'              2. CORE DATA POPULATION & RESTORATION SUB-ROUTINES
'================================================================================

'------------------------------------------------------------------------------
' PopulateMasterQuotesData - Fills Dashboard A-L from MasterQuotes_Final Source
' Handles Table or Named Range source. Populates K from [AutoNote], L from [AutoStage].
' *** Using robust version ***
'------------------------------------------------------------------------------
Private Sub PopulateMasterQuotesData(ws As Worksheet)
    If ws Is Nothing Then Exit Sub

    Dim sourceNameInput As String: sourceNameInput = MASTER_QUOTES_FINAL_SOURCE
    Dim sourceReference As String ' How source is referenced in formulas (Table Name or Sheet!Range)
    Dim isTable As Boolean: isTable = False
    Dim isNamedRange As Boolean: isNamedRange = False
    Dim sourceHeaderRow As Long
    Dim dataStartRowOffset As Long ' Usually 1 for Tables/Named Ranges with headers
    Dim colMap As Object ' Dictionary to map header names -> column index for Named Ranges
    Dim lastMasterRow As Long ' Number of data rows in the source
    Dim targetRowCount As Long ' Number of rows to populate on dashboard

    Dim PopulateErrorHandler As String ' Declare error handler label variable
    PopulateErrorHandler = "PopulateErrorHandler" ' Assign string name

    On Error GoTo PopulateErrorHandler_Handler ' Use named error handler

    Dim sourceTable As ListObject
    Dim sourceRange As Range

    ' --- Determine Source Type and Validate ---
    Set sourceTable = Nothing: On Error Resume Next
    Dim tempWs As Worksheet
    For Each tempWs In ThisWorkbook.Worksheets
        Set sourceTable = tempWs.ListObjects(sourceNameInput)
        If Not sourceTable Is Nothing Then Exit For
    Next tempWs
    isTable = Not sourceTable Is Nothing
    On Error GoTo PopulateErrorHandler_Handler ' Restore error handler

    If isTable Then
        ' --- Source is a TABLE ---
        LogUserEditsOperation "PopulateMasterQuotesData: Source '" & sourceNameInput & "' found as a Table on sheet '" & sourceTable.Parent.Name & "'."
        If sourceTable.DataBodyRange Is Nothing Then
             LogUserEditsOperation "PopulateMasterQuotesData: Source Table has no data rows."
             GoTo SourceEmpty_Handler ' Use named handler
        End If
        On Error Resume Next ' Check row count safely
        lastMasterRow = sourceTable.ListRows.Count
        If Err.Number <> 0 Or lastMasterRow = 0 Then Err.Clear: GoTo SourceEmpty_Handler ' Use named handler
        On Error GoTo PopulateErrorHandler_Handler ' Restore error handler
        sourceReference = sourceNameInput ' Use Table name directly in INDEX formula
        sourceHeaderRow = sourceTable.HeaderRowRange.Row
        dataStartRowOffset = 1 ' Data starts 1 row below header in a table
    Else
        ' --- Check if source is a NAMED RANGE ---
        Set sourceRange = Nothing: On Error Resume Next
        Set sourceRange = ThisWorkbook.Names(sourceNameInput).RefersToRange
        isNamedRange = (Err.Number = 0 And Not sourceRange Is Nothing)
        On Error GoTo PopulateErrorHandler_Handler ' Restore error handler

        If isNamedRange Then
            LogUserEditsOperation "PopulateMasterQuotesData: Source '" & sourceNameInput & "' found as a Named Range on sheet '" & sourceRange.Parent.Name & "'."
            ' Assume header is in the first row of the named range
            sourceHeaderRow = sourceRange.Row
            dataStartRowOffset = 1
            ' Count non-blank cells in first column to estimate rows (including header)
            On Error Resume Next
            Dim totalRowsInRange As Long
            totalRowsInRange = Application.WorksheetFunction.CountA(sourceRange.Columns(1))
             If Err.Number <> 0 Or totalRowsInRange <= dataStartRowOffset Then Err.Clear: GoTo SourceEmpty_Handler ' Use named handler
            On Error GoTo PopulateErrorHandler_Handler ' Restore error handler
            lastMasterRow = totalRowsInRange - dataStartRowOffset ' Actual data rows = Total - Header rows

            ' Use qualified Sheet!Range address for INDEX with Named Ranges
            sourceReference = "'" & sourceRange.Worksheet.Name & "'!" & sourceRange.Address(True, True)

            ' Build header map (Header Name -> Column Index) for the Named Range
            Set colMap = CreateObject("Scripting.Dictionary")
            colMap.CompareMode = vbTextCompare ' Case-insensitive header matching
            Dim headerData As Variant: headerData = sourceRange.rows(1).Value
            Dim c As Long
            If IsArray(headerData) Then
                For c = 1 To UBound(headerData, 2)
                    If Not IsEmpty(headerData(1, c)) Then colMap(CStr(headerData(1, c))) = c
                Next c
            ElseIf Not IsEmpty(headerData) Then ' Single column range
                colMap(CStr(headerData)) = 1
            End If
            LogUserEditsOperation "PopulateMasterQuotesData: Built header map for Named Range. Found " & colMap.Count & " columns."
            ' Check if essential columns exist in the map
            ' *** Relaxed check - will handle missing columns in loop ***
            If Not colMap.Exists("Document Number") Then
                 LogUserEditsOperation "PopulateMasterQuotesData: ERROR - Required column 'Document Number' not found in Named Range headers."
                 GoTo SourceInvalid_Handler ' Use named handler
            End If
        Else
            ' --- Source Not Found or Invalid ---
             LogUserEditsOperation "PopulateMasterQuotesData: ERROR - Source '" & sourceNameInput & "' not found as a Table or a valid Named Range."
             GoTo SourceInvalid_Handler ' Use named handler
        End If
    End If

    ' --- Final Row Count Check ---
    targetRowCount = lastMasterRow
    If targetRowCount <= 0 Then GoTo SourceEmpty_Handler ' Use named handler
    LogUserEditsOperation "PopulateMasterQuotesData: Source contains " & targetRowCount & " data rows to populate."

    ' --- Populate Dashboard A-L with Formulas ---
    Application.Calculation = xlCalculationManual ' Ensure calc is off

    Dim colHeadersNeeded As Variant
    colHeadersNeeded = Array("Document Number", "Customer Number", "Customer Name", "Document Amount", _
                           "Document Date", "First Date Pulled", "Salesperson ID", "User To Enter", _
                           "Pull Count", "AutoNote", "AutoStage", "DataSource") ' Include DataSource for Col L logic
    Dim dashColsTarget As Variant
    dashColsTarget = Array("A", "B", "C", "D", "E", "F", "G", "H", "I", DB_COL_MISSING_QUOTE, DB_COL_PHASE) ' A-I, K, L
    Dim isNumeric As Variant ' Flags for --INDEX conversion
    isNumeric = Array(False, False, False, True, True, True, False, False, True, False, False) ' D,E,F,I are numeric/date

    Dim i As Long, formula As String, sourceColName As String, targetCol As String
    Dim colRef As String, colNum As Long

    With ws
        For i = LBound(colHeadersNeeded) To UBound(colHeadersNeeded) - 1 ' Loop through needed headers (excluding DataSource)
             sourceColName = colHeadersNeeded(i)
             targetCol = dashColsTarget(i) ' Corresponding dashboard column A-I, K, L
             formula = "" ' Reset formula for each column

             ' --- Determine Column Reference Syntax (Table vs Named Range) & Check Existence ---
             If isTable Then
                 On Error Resume Next
                 Dim listCol As ListColumn: Set listCol = Nothing
                 Set listCol = sourceTable.ListColumns(sourceColName)
                 If Err.Number <> 0 Or listCol Is Nothing Then
                     LogUserEditsOperation "PopulateMasterQuotesData: WARNING - Column '" & sourceColName & "' not found in Table source. Skipping dashboard column " & targetCol & "."
                     Err.Clear
                     GoTo NextColumn_Handler ' Skip if column doesn't exist
                 End If
                 On Error GoTo PopulateErrorHandler_Handler ' Restore handler
                 colRef = "[" & sourceColName & "]" ' Table syntax [Column Name]
             Else ' isNamedRange
                 If colMap.Exists(sourceColName) Then
                     colNum = colMap(sourceColName)
                     colRef = "," & colNum ' Named Range syntax uses ,ColNum for INDEX
                 Else
                     LogUserEditsOperation "PopulateMasterQuotesData: WARNING - Column '" & sourceColName & "' not found in Named Range source headers. Skipping dashboard column " & targetCol & "."
                     GoTo NextColumn_Handler ' Skip if header not found
                 End If
             End If

             ' --- Construct Base INDEX Formula ---
             If isTable Then
                 formula = "=IFERROR(INDEX(" & sourceReference & colRef & ",ROWS($A$4:A4)),"""")"
             Else ' isNamedRange
                 formula = "=IFERROR(INDEX(" & sourceReference & ",ROWS($A$4:A4)+" & dataStartRowOffset - 1 & colRef & "),"""")"
             End If

             ' --- Apply Special Logic / Formatting ---
             If isNumeric(i) Then formula = Replace$(formula, "INDEX", "--INDEX")

             If targetCol = DB_COL_PHASE Then ' Special logic for Column L (Phase)
                 Dim dataSourceColRef As String: dataSourceColRef = ""
                 Dim dsColNum As Long: dsColNum = 0
                 Dim listColDS As ListColumn: Set listColDS = Nothing

                 If isTable Then
                     On Error Resume Next
                     Set listColDS = sourceTable.ListColumns("DataSource")
                     If Err.Number = 0 And Not listColDS Is Nothing Then dataSourceColRef = "[DataSource]"
                     Err.Clear
                     On Error GoTo PopulateErrorHandler_Handler
                 Else ' isNamedRange
                     If colMap.Exists("DataSource") Then
                        dsColNum = colMap("DataSource")
                        dataSourceColRef = ",ROWS($A$4:A4)+" & dataStartRowOffset - 1 & "," & dsColNum
                     End If
                 End If

                 If dataSourceColRef <> "" Or dsColNum > 0 Then
                    Dim dataSourceIndexPart As String
                    If isTable Then
                        dataSourceIndexPart = "INDEX(" & sourceReference & dataSourceColRef & ",ROWS($A$4:A4))"
                    Else ' isNamedRange
                        dataSourceIndexPart = "INDEX(" & sourceReference & dataSourceColRef & ")"
                    End If
                    formula = "=IF(" & dataSourceIndexPart & "=""LEGACY""," & Mid$(formula, 2) & ","""")" ' Insert IF condition
                 Else
                      LogUserEditsOperation "PopulateMasterQuotesData: WARNING - 'DataSource' column needed for Column L logic not found. Column L formula cleared."
                      formula = "" ' Clear formula if DataSource is missing
                 End If
             End If

             ' Add outer IF(A4<>"",...) wrapper for columns B onwards
             If targetCol <> "A" And formula <> "" Then
                  formula = "=IF(A4<>""""," & Mid$(formula, 2) & ","""")"
             End If

             ' --- Write Formula to Target Column ---
             If formula <> "" Then
                 On Error Resume Next
                 .Range(targetCol & "4").Resize(targetRowCount, 1).formula = formula
                 If Err.Number <> 0 Then
                     LogUserEditsOperation "ERROR writing formula for column " & targetCol & ": " & Err.Description & " Formula: " & formula
                     Err.Clear
                 End If
                 On Error GoTo PopulateErrorHandler_Handler
             End If

NextColumn_Handler: ' Label to jump to for skipping a column
        Next i

        ' --- Apply Basic Formatting ---
        On Error Resume Next
        .Range("D4").Resize(targetRowCount).NumberFormat = "$#,##0.00"   ' Amount (D)
        .Range("E4").Resize(targetRowCount).NumberFormat = "mm/dd/yyyy" ' Document Date (E)
        .Range("F4").Resize(targetRowCount).NumberFormat = "mm/dd/yyyy" ' First Date Pulled (F)
        .Range("I4").Resize(targetRowCount).NumberFormat = "0"           ' Pull Count (I) - Integer
        On Error GoTo PopulateErrorHandler_Handler ' Use named handler

    End With ' ws

    LogUserEditsOperation "PopulateMasterQuotesData: Finished applying formulas for A-L."
    GoTo Cleanup_Populate ' Go to cleanup

SourceEmpty_Handler:
    LogUserEditsOperation "PopulateMasterQuotesData: Source '" & sourceNameInput & "' is empty. No data populated for A-L."
    GoTo Cleanup_Populate
SourceInvalid_Handler:
    LogUserEditsOperation "PopulateMasterQuotesData: Source '" & sourceNameInput & "' is invalid or missing required columns. No data populated for A-L."
    GoTo Cleanup_Populate
PopulateErrorHandler_Handler:
    LogUserEditsOperation "ERROR in PopulateMasterQuotesData: [" & Err.Number & "] " & Err.Description & " (Line: " & Erl & ")"
    ' Decide if error is fatal - maybe re-raise or allow RefreshDashboard handler to catch
Cleanup_Populate:
    On Error Resume Next ' Prevent error during cleanup
    Set colMap = Nothing: Set sourceTable = Nothing: Set sourceRange = Nothing
    Set listCol = Nothing: Set listColDS = Nothing: Set tempWs = Nothing
    Set headerData = Nothing ' Release variant array
    ' Resume Next ' Uncomment this if you want execution to continue in RefreshDashboard after an error here
End Sub

'------------------------------------------------------------------------------
' PopulateWorkflowLocation – Fills Dashboard J from DocNum_LatestLocation PQ
' Silent unless DEBUG_WorkflowLocation = True
' *** Using robust version ***
'------------------------------------------------------------------------------
Private Sub PopulateWorkflowLocation(ws As Worksheet, lastRowDash As Long)
    If lastRowDash < 4 Then Exit Sub ' No data rows on dashboard

    Dim wsPQ As Worksheet, tbl As ListObject
    Dim dict As Object ' Dictionary: CleanedDocNum -> Location
    Dim arr As Variant, pqDataRange As Range
    Dim r As Long, i As Long
    Dim colDoc As Long, colLoc As Long
    Dim key As String, loc As String
    Dim t_Start As Double: t_Start = Timer
    Dim foundCount As Long, notFoundCount As Long

    LogUserEditsOperation "PopulateWorkflowLocation: Starting update for Column J."

    ' --- Get Source Table ---
    On Error Resume Next
    Set wsPQ = Nothing: Set tbl = Nothing
    Set wsPQ = ThisWorkbook.Sheets(PQ_LATEST_LOCATION_SHEET)
    If wsPQ Is Nothing Then
        LogUserEditsOperation "PopulateWorkflowLocation: ERROR - Source sheet '" & PQ_LATEST_LOCATION_SHEET & "' not found."
        Exit Sub
    End If
    Set tbl = wsPQ.ListObjects(PQ_LATEST_LOCATION_TABLE)
    If tbl Is Nothing Then Set tbl = wsPQ.ListObjects(1) ' Fallback to first table
    On Error GoTo 0 ' Restore default error handling

    If tbl Is Nothing Then
        LogUserEditsOperation "PopulateWorkflowLocation: ERROR - Source table '" & PQ_LATEST_LOCATION_TABLE & "' (or first table) not found on sheet '" & wsPQ.Name & "'."
        Exit Sub
    End If
    LogUserEditsOperation "PopulateWorkflowLocation: Found source table '" & tbl.Name & "' on sheet '" & wsPQ.Name & "'."

    ' --- Build Lookup Dictionary ---
    Set dict = CreateObject("Scripting.Dictionary")
    dict.CompareMode = vbTextCompare ' Case-insensitive keys

    On Error Resume Next ' Handle errors reading table data / finding columns
    Set pqDataRange = tbl.DataBodyRange
    If pqDataRange Is Nothing Then
        LogUserEditsOperation "PopulateWorkflowLocation: Source table '" & tbl.Name & "' has no data rows."
    Else
        arr = pqDataRange.Value ' Read data into array for speed
        colDoc = 0: colLoc = 0
        colDoc = tbl.ListColumns(PQ_DOCNUM_COL_NAME).Index      ' Get index using constant
        colLoc = tbl.ListColumns(PQ_LOCATION_COL_NAME).Index    ' Get index using constant
        If Err.Number <> 0 Or colDoc = 0 Or colLoc = 0 Then
             LogUserEditsOperation "PopulateWorkflowLocation: ERROR - Required columns ('" & PQ_DOCNUM_COL_NAME & "' or '" & PQ_LOCATION_COL_NAME & "') not found in source table '" & tbl.Name & "'."
             GoTo Cleanup_Workflow ' Go to cleanup on error
        End If

        ' Populate dictionary using CLEANED keys
        For r = 1 To UBound(arr, 1)
            key = CleanDocumentNumber(CStr(arr(r, colDoc))) ' Clean the source DocNum
            If Len(key) > 0 Then
                loc = CStr(arr(r, colLoc)) ' Get location value
                If Not dict.Exists(key) Then dict.Add key, loc
            End If
        Next r
        LogUserEditsOperation "PopulateWorkflowLocation: Built dictionary with " & dict.Count & " entries from source table."
    End If
    On Error GoTo WorkflowErrorHandler ' Use specific handler from here

    ' --- Write Locations to Dashboard Column J ---
    Dim dashDocNum As String
    Dim workflowLoc As String
    Dim debugMsg As String ' For optional debug logging

    If dict.Count > 0 Then ' Only loop if dictionary has entries
        For i = 4 To lastRowDash
            dashDocNum = CleanDocumentNumber(CStr(ws.Cells(i, "A").Value)) ' Clean dashboard DocNum
            workflowLoc = "Quote Only" ' Default value

            If Len(dashDocNum) > 0 Then
                If dict.Exists(dashDocNum) Then
                    workflowLoc = CStr(dict(dashDocNum)) ' Get location from dictionary
                    ' Ensure blank/null locations default to "Quote Only"
                    If Len(Trim(workflowLoc)) = 0 Then workflowLoc = "Quote Only"
                    foundCount = foundCount + 1
                Else
                    notFoundCount = notFoundCount + 1 ' DocNum on dashboard not found in location source
                End If
            End If

            ws.Cells(i, DB_COL_WORKFLOW_LOCATION).Value = workflowLoc ' Write value to Column J

            ' Optional Debug Logging (controlled by constant)
            If DEBUG_WorkflowLocation Then
                debugMsg = "WorkflowLoc row " & i & ": DashKey='" & dashDocNum & "'" & _
                           IIf(dict.Exists(dashDocNum), " FOUND -> '", " NOT FOUND -> Defaulting to '") & workflowLoc & "'"
                Debug.Print debugMsg ' Output to Immediate Window
            End If
        Next i
        LogUserEditsOperation "PopulateWorkflowLocation: Finished writing to Column J. Found: " & foundCount & ", Not Found: " & notFoundCount & "."
    Else
        LogUserEditsOperation "PopulateWorkflowLocation: Lookup dictionary is empty. Writing default 'Quote Only' to Column J."
        ' If dictionary is empty, fill column J with default
        If lastRowDash >= 4 Then ws.Range(DB_COL_WORKFLOW_LOCATION & "4:" & DB_COL_WORKFLOW_LOCATION & lastRowDash).Value = "Quote Only"
    End If

    LogUserEditsOperation "PopulateWorkflowLocation: Completed. Time: " & Format(Timer - t_Start, "0.00") & "s"

Cleanup_Workflow: ' Cleanup point for this sub
    On Error Resume Next ' Prevent errors during cleanup
    Set dict = Nothing: Set tbl = Nothing: Set wsPQ = Nothing: Set pqDataRange = Nothing
    If IsArray(arr) Then Erase arr
    Exit Sub

WorkflowErrorHandler: ' Error handler for this sub
    LogUserEditsOperation "ERROR in PopulateWorkflowLocation: [" & Err.Number & "] " & Err.Description & " (Line: " & Erl & ")"
    Resume Cleanup_Workflow ' Go to cleanup on error

End Sub


'------------------------------------------------------------------------------
' RestoreUserEditsToDashboard – Applies UserEdits B,C,D -> Dashboard L,M,N
' Uses robust dictionary lookup with cleaned keys.
' *** Using robust version ***
'------------------------------------------------------------------------------
Private Sub RestoreUserEditsToDashboard(wsDash As Worksheet, _
                                        wsEdits As Worksheet, _
                                        lastRowDash As Long)

    LogUserEditsOperation "RestoreUserEdits: Starting process..."
    If wsDash Is Nothing Or wsEdits Is Nothing Then
        LogUserEditsOperation "RestoreUserEdits: ERROR - Worksheet object(s) not provided."
        Exit Sub
    End If
    If lastRowDash < 4 Then
        LogUserEditsOperation "RestoreUserEdits: No data rows on dashboard to restore edits to."
        Exit Sub
    End If

    ' --- Load UserEdits dictionary with cleaned keys ---
    Dim dict As Object
    Set dict = LoadUserEditsToDictionary(wsEdits) ' Assumes this function handles errors and logs
    If dict Is Nothing Then ' Check if dictionary creation failed critically
         LogUserEditsOperation "RestoreUserEdits: CRITICAL ERROR - Failed to load UserEdits dictionary. Aborting restore."
         Exit Sub
    End If

    If dict.Count = 0 Then
        LogUserEditsOperation "RestoreUserEdits: No data found in UserEdits dictionary. Skipping restore."
        Exit Sub
    End If

    Dim rDash As Long, rEdit As Long
    Dim keyDash As String
    Dim restoredCount As Long, notFoundCount As Long
    Dim errCount As Long

    LogUserEditsOperation "RestoreUserEdits: Processing dashboard rows 4 to " & lastRowDash & "..."

    ' --- Loop through Dashboard Rows ---
    For rDash = 4 To lastRowDash
        keyDash = CleanDocumentNumber(CStr(wsDash.Cells(rDash, "A").Value)) ' Clean the dashboard key

        If Len(keyDash) > 0 Then
            If dict.Exists(keyDash) Then
                ' --- Match Found: Restore Data ---
                rEdit = dict(keyDash) ' Get the SHEET row number from UserEdits
                restoredCount = restoredCount + 1

                On Error Resume Next ' Handle potential errors writing to dashboard cell for THIS ROW ONLY
                wsDash.Cells(rDash, DB_COL_PHASE).Value = wsEdits.Cells(rEdit, UE_COL_PHASE).Value
                wsDash.Cells(rDash, DB_COL_LASTCONTACT).Value = wsEdits.Cells(rEdit, UE_COL_LASTCONTACT).Value
                wsDash.Cells(rDash, DB_COL_COMMENTS).Value = wsEdits.Cells(rEdit, UE_COL_COMMENTS).Value
                If Err.Number <> 0 Then
                    errCount = errCount + 1
                    ' Log only the first few errors to avoid flooding log
                    If errCount <= 5 Then LogUserEditsOperation "RestoreUserEdits: ERROR writing data for Dash Row " & rDash & " (Key: " & keyDash & ") from UE Row " & rEdit & ". Error: " & Err.Description
                    Err.Clear
                End If
                On Error GoTo 0 ' Restore default error handling for the loop

            Else
                ' --- No Match Found ---
                ' Dashboard DocNum exists but has no corresponding entry in UserEdits.
                ' Keep the value in Col L (from AutoStage). Clear M and N.
                wsDash.Cells(rDash, DB_COL_LASTCONTACT).ClearContents ' Clear M
                wsDash.Cells(rDash, DB_COL_COMMENTS).ClearContents    ' Clear N
                notFoundCount = notFoundCount + 1
            End If
        Else
            ' Dashboard key is blank, clear L-N
            wsDash.Range(wsDash.Cells(rDash, DB_COL_PHASE), wsDash.Cells(rDash, DB_COL_COMMENTS)).ClearContents
        End If
    Next rDash

    If errCount > 0 Then LogUserEditsOperation "RestoreUserEdits: Encountered " & errCount & " errors writing data (logged first 5)."
    LogUserEditsOperation "RestoreUserEdits: Finished. Edits applied/checked: " & restoredCount & ". No matching edit found: " & notFoundCount & "."
    Set dict = Nothing ' Clean up

End Sub


'================================================================================
'              3. LAYOUT, SETUP, and UI ROUTINES
'              *** Using User's Preferred Versions ***
'================================================================================

'------------------------------------------------------------------------------
' GetOrCreateDashboardSheet - Returns the main dashboard Worksheet object
' *** Uses user's preferred version which calls SetupDashboard ***
'------------------------------------------------------------------------------
Private Function GetOrCreateDashboardSheet(sheetName As String) As Worksheet
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(sheetName)
    On Error GoTo 0

    If ws Is Nothing Then
        LogUserEditsOperation "Dashboard sheet '" & sheetName & "' not found. Creating new sheet." ' Added log
        Set ws = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        ws.Name = sheetName
        ' Call user's SetupDashboard to establish initial layout for new sheet
        SetupDashboard ws ' Call the user's specific SetupDashboard for layout
        LogUserEditsOperation "Called SetupDashboard for new sheet." ' Added log
    End If

    Set GetOrCreateDashboardSheet = ws
End Function


'------------------------------------------------------------------------------
' CleanupDashboardLayout - Clears old data rows (4+), preserves rows 1-3
' *** Uses user's provided version ***
'------------------------------------------------------------------------------
Private Sub CleanupDashboardLayout(ws As Worksheet)
    Application.ScreenUpdating = False
    LogUserEditsOperation "CleanupDashboardLayout: Clearing data rows 4+ on sheet '" & ws.Name & "'." ' Added Log

    Dim lastRow As Long
    lastRow = ws.Cells(ws.rows.Count, "A").End(xlUp).Row
    Dim dataRange As Range
    Dim tempData As Variant

    ' This section seems intended to preserve data during cleanup,
    ' but the actual .Clear command later removes it anyway.
    ' Keeping the logic as provided by user, but noting it might be redundant.
    ' If lastRow >= 4 Then
    '     Set dataRange = ws.Range("A4:" & DB_COL_COMMENTS & lastRow)
    '     tempData = dataRange.Value
    ' End If

    ' This check seems unnecessary if InitializeDashboardLayout sets headers anyway
    ' Dim hasTitle As Boolean
    ' hasTitle = False
    ' Dim cell As Range
    ' For Each cell In ws.Range("A1:" & DB_COL_COMMENTS & "1").Cells
    '     If InStr(1, CStr(cell.Value), "STRATEGIC QUOTE RECOVERY", vbTextCompare) > 0 Then
    '         hasTitle = True
    '         Exit For
    '     End If
    ' Next cell

    ' --- The Actual Clearing Action ---
    On Error Resume Next ' Ignore error if already clear or protected
    ws.Range("A4:" & DB_COL_COMMENTS & ws.rows.Count).ClearContents ' Clear Data rows only
    If Err.Number <> 0 Then LogUserEditsOperation "CleanupDashboardLayout: Note - Error during clear contents.": Err.Clear
    On Error GoTo 0
    ' --- End Clearing ---

    ' This section recreating rows 1-3 seems redundant if InitializeDashboardLayout does it.
    ' Commenting out to avoid potential conflicts, relying on InitializeDashboardLayout.
    ' If Not hasTitle Then
    '      With ws.Range("A1:" & DB_COL_COMMENTS & "1")
    '          ... (Title formatting) ...
    '      End With
    ' End If
    ' With ws.Range("A2:" & DB_COL_COMMENTS & "2")
    '      ... (Row 2 formatting) ...
    ' End With
    ' With ws.Range("A2")
    '      ... (Control panel label formatting) ...
    ' End With
    ' With ws.Range(DB_COL_COMMENTS & "2")
    '      ... (? formatting) ...
    ' End With
    ' With ws.Range("A3:" & DB_COL_COMMENTS & "3")
    '      ... (Header formatting) ...
    ' End With

    ' This restore logic also seems misplaced here if RefreshDashboard handles restore
    ' If Not IsEmpty(tempData) Then
    '      ... (Restore tempData) ...
    ' End If

    Application.ScreenUpdating = True
End Sub


'------------------------------------------------------------------------------
' InitializeDashboardLayout - Sets headers A-N in Row 3
' *** MODIFIED to enable wrap text for Comments column N ***
'------------------------------------------------------------------------------
Private Sub InitializeDashboardLayout(ws As Worksheet)
    If ws Is Nothing Then Exit Sub
    LogUserEditsOperation "InitializeDashboardLayout: Setting headers A-N (Using AutoFit later)."

    ' --- Clear Data Rows & Extra Columns ---
    ws.Range("A4:" & DB_COL_COMMENTS & ws.rows.Count).ClearContents
    On Error Resume Next ' Ignore error deleting columns
    ws.Range("O:" & ws.Columns.Count).Delete Shift:=xlToLeft
    On Error GoTo 0 ' Restore default error handling for this sub

    ' --- Set Headers in Row 3 ---
    With ws.Range("A3:" & DB_COL_COMMENTS & "3") ' Range A3:N3
        .ClearContents
        .Value = Array( _
            "Document Number", "Client ID", "Customer Name", "Document Amount", "Document Date", _
            "First Date Pulled", "Salesperson ID", "Entered By", "Pull Count", "Workflow Location", _
            "Missing Quote Alert", "Engagement Phase", "Last Contact Date", "User Comments")
        .Font.Bold = True
        .Interior.Color = RGB(16, 107, 193)
        .Font.Color = RGB(255, 255, 255)
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = False ' Headers should NOT wrap
        .Borders(xlEdgeBottom).LineStyle = xlContinuous
        .Borders(xlEdgeBottom).Weight = xlHairline
        .Borders(xlEdgeBottom).Color = RGB(200, 200, 200)
    End With

    ' --- REMOVED All Fixed Column Width Settings ---

    ' --- AutoFit Row 3 Height ---
    On Error Resume Next
    ws.rows(3).AutoFit
    On Error GoTo 0

    ' *** ADDED: Enable Text Wrapping for the data area of Column N ***
    On Error Resume Next ' Ignore error if sheet is protected
    ' Apply to a reasonable number of rows, or adjust as needed
    Dim wrapRange As Range
    Set wrapRange = ws.Range(DB_COL_COMMENTS & "4:" & DB_COL_COMMENTS & ws.rows.Count) ' Select Col N data area
    wrapRange.WrapText = True
    Set wrapRange = Nothing ' Clear object variable
    If Err.Number <> 0 Then LogUserEditsOperation "InitializeDashboardLayout: Warning - could not set WrapText for Column N.": Err.Clear
    On Error GoTo 0 ' Use specific handler if defined for this sub, else default

    LogUserEditsOperation "InitializeDashboardLayout: Headers set (Column widths to be AutoFit later)."
End Sub


'------------------------------------------------------------------------------
' SetupDashboard - Sets up static Rows 1 (Title) and 2 (Control Panel)
' *** This appears to be the user's preferred setup for Rows 1 & 2 ***
'------------------------------------------------------------------------------
Public Sub SetupDashboard(ws As Worksheet)
     LogUserEditsOperation "SetupDashboard: Setting up Title (Row 1) and Control Panel (Row 2)." ' Added Log
     Application.ScreenUpdating = False
     On Error Resume Next ' Ignore errors if sheet is protected

     ' --- Row 1: Title Bar ---
     With ws.Range("A1:" & DB_COL_COMMENTS & "1") ' A1:N1
         .ClearContents ' Clear previous content/merge
         .Merge
         .Value = "STRATEGIC QUOTE RECOVERY & CONVERSION TRACKER"
         .HorizontalAlignment = xlCenter
         .VerticalAlignment = xlCenter
         .Font.Size = 18
         .Font.Bold = True
         .Interior.Color = RGB(16, 107, 193) ' Blue background
         .Font.Color = RGB(255, 255, 255) ' White text
         .RowHeight = 32
     End With

     ' --- Row 2: Control Panel Area ---
     With ws.Range("A2:" & DB_COL_COMMENTS & "2") ' A2:N2
         .ClearContents
         .Interior.Color = RGB(245, 245, 245) ' Light grey background
         .Borders(xlEdgeTop).LineStyle = xlContinuous
         .Borders(xlEdgeTop).Weight = xlThin
         .Borders(xlEdgeTop).Color = RGB(200, 200, 200)
         .Borders(xlEdgeBottom).LineStyle = xlContinuous
         .Borders(xlEdgeBottom).Weight = xlThin
         .Borders(xlEdgeBottom).Color = RGB(200, 200, 200)
         .RowHeight = 28
         .VerticalAlignment = xlCenter
     End With

     ' --- Row 2: "CONTROL PANEL" Label (User preferred A2 only) ---
     With ws.Range("A2")
         .Value = "CONTROL PANEL"
         .Font.Bold = True
         .Font.Size = 10
         .Font.Name = "Segoe UI" ' Or original font
         .HorizontalAlignment = xlCenter
         .VerticalAlignment = xlCenter
         .Interior.Color = RGB(70, 130, 180) ' Steel blue
         .Font.Color = RGB(255, 255, 255)
         .ColumnWidth = 16 ' From user's code
         .Borders(xlEdgeRight).LineStyle = xlContinuous
         .Borders(xlEdgeRight).Weight = xlThin
         .Borders(xlEdgeRight).Color = RGB(200, 200, 200)
     End With
     ' Clear B2 if A2 is used and merging is not desired
      ws.Range("B2").ClearContents

      ' --- Row 2: Placeholder for Timestamp (Merged G2:I2) ---
      ws.Range("G2:I2").Merge ' Ensure merged

      ' --- Row 2: Help (?) Icon ---
      With ws.Range(DB_COL_COMMENTS & "2") ' N2
          .Value = "?"
          .Font.Bold = True
          .Font.Size = 14
          .HorizontalAlignment = xlCenter
          .VerticalAlignment = xlCenter
          .Font.Color = RGB(70, 130, 180) ' Match steel blue
      End With

      ' --- Buttons are created later by SetupDashboardUI_EndRefresh ---

     If Err.Number <> 0 Then LogUserEditsOperation "SetupDashboard: Note - Error setting up rows 1-2.": Err.Clear
     On Error GoTo 0
     Application.ScreenUpdating = True
 End Sub


'------------------------------------------------------------------------------
' ModernButton - Creates styled buttons
' *** Using compatible version (avoids TextFrame2 errors) ***
'------------------------------------------------------------------------------
Public Sub ModernButton(ws As Worksheet, cellRef As String, buttonText As String, macroName As String)
    Dim btn As Shape
    Dim targetCell As Range
    Dim btnLeft As Double, btnTop As Double, btnWidth As Double, btnHeight As Double

    On Error Resume Next ' Handle invalid cellRef
    Set targetCell = ws.Range(cellRef)
    If targetCell Is Nothing Then
        LogUserEditsOperation "ModernButton: ERROR - Invalid cell reference '" & cellRef & "'."
        Exit Sub
    End If
    On Error GoTo 0 ' Restore error handling

    ' Calculate button position and size based on target cell(s)
    ' Use user's preferred sizing logic
    btnLeft = targetCell.Left
    btnTop = targetCell.Top
    btnWidth = targetCell.Width * 1.6
    btnHeight = targetCell.Height * 0.75
    btnTop = btnTop + (targetCell.Height - btnHeight) / 2 ' Center vertically

    ' Add the shape
    On Error Resume Next ' Handle error adding shape
    Set btn = ws.Shapes.AddShape(msoShapeRoundedRectangle, btnLeft, btnTop, btnWidth, btnHeight)
    If Err.Number <> 0 Then
        LogUserEditsOperation "ModernButton: ERROR creating button shape '" & buttonText & "'. Error: " & Err.Description
        Set btn = Nothing ' Ensure btn is Nothing on failure
        Exit Sub
    End If
    On Error GoTo 0

    ' Style the button
    With btn
        .Fill.Visible = msoTrue
        .Fill.Solid
        .Fill.ForeColor.RGB = RGB(42, 120, 180) ' User's preferred blue
        .Line.ForeColor.RGB = RGB(25, 95, 150)  ' Darker blue line
        .Line.Weight = 0.75
        .Name = "btn" & Replace(buttonText, " ", "") ' Give button a name

        ' Text Formatting using reliable TextFrame
        On Error Resume Next
        With .TextFrame
            .Characters.Text = buttonText
            .Characters.Font.Color = RGB(255, 255, 255) ' White text
            .Characters.Font.Size = 10
            .Characters.Font.Name = "Segoe UI" ' Match user's preference
            .Characters.Font.Bold = True
            .HorizontalAlignment = xlHAlignCenter
            .VerticalAlignment = xlVAlignCenter
        End With
        If Err.Number <> 0 Then LogUserEditsOperation "ModernButton: Warning - Error setting text format for '" & buttonText & "'.": Err.Clear

        ' Shadow Effect (Keep user's preferred shadow if it worked)
        On Error Resume Next
        .Shadow.Type = msoShadow21 ' Check if this is the desired one
        .Shadow.Transparency = 0.7
        .Shadow.Visible = msoTrue
         If Err.Number <> 0 Then LogUserEditsOperation "ModernButton: Warning - Error setting shadow for '" & buttonText & "'.": Err.Clear

        ' Assign Macro
        .OnAction = macroName
         If Err.Number <> 0 Then LogUserEditsOperation "ModernButton: Error assigning macro '" & macroName & "'.": Err.Clear

    End With
    On Error GoTo 0
    Set btn = Nothing: Set targetCell = Nothing
End Sub


'------------------------------------------------------------------------------
' FreezeDashboard - Freezes rows 1-3
' *** Uses user's provided version ***
'------------------------------------------------------------------------------
Private Sub FreezeDashboard(ws As Worksheet)
    If ws Is Nothing Then Exit Sub ' Added safety check
    LogUserEditsOperation "FreezeDashboard: Freezing panes at row 4." ' Added Log
    ws.Activate ' Ensure sheet is active
    On Error Resume Next ' Ignore error if already frozen/unfrozen
    ActiveWindow.FreezePanes = False ' Unfreeze first
    ws.Range("A4").Select          ' Select cell below freeze row
    ActiveWindow.FreezePanes = True  ' Freeze above selected cell
    ws.Range("A1").Select ' Select A1 after freezing
    If Err.Number <> 0 Then LogUserEditsOperation "FreezeDashboard: Note - Error during freeze panes operation.": Err.Clear
    On Error GoTo 0
End Sub

'------------------------------------------------------------------------------
' SetupDashboardUI_EndRefresh - Adds Timestamp and Buttons to Row 2 at end of refresh
' *** NEW Sub to consolidate UI updates called at end of RefreshDashboard ***
'------------------------------------------------------------------------------
Private Sub SetupDashboardUI_EndRefresh(ws As Worksheet)
    If ws Is Nothing Then Exit Sub
    LogUserEditsOperation "SetupDashboardUI_EndRefresh: Updating Timestamp and Buttons."
    On Error Resume Next ' Ignore errors if sheet is protected

    ' --- Update Timestamp (Merged G2:I2) ---
    With ws.Range("G2:I2") ' Merged by Initialize or SetupDashboardLayout
        .Value = "Last Refreshed: " & Format$(Now(), "mm/dd/yyyy h:nn AM/PM")
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .Font.Size = 9
        .Font.Name = "Segoe UI" ' Match user preference
        .Font.Color = RGB(80, 80, 80) ' Dark gray
    End With
    If Err.Number <> 0 Then LogUserEditsOperation "SetupDashboardUI_EndRefresh: Error setting timestamp.": Err.Clear

    ' --- Remove Old Buttons in Row 2 ---
    Dim shp As Shape
    For Each shp In ws.Shapes
        If shp.TopLeftCell.Row = 2 Then
            shp.Delete
        End If
    Next shp
    If Err.Number <> 0 Then LogUserEditsOperation "SetupDashboardUI_EndRefresh: Error deleting old buttons.": Err.Clear

    ' --- Create New Buttons using user's ModernButton ---
    ' Create buttons in single cells C2 and E2 as per user's preference
    ModernButton ws, "C2", "Standard Refresh", "Button_RefreshDashboard_SaveAndRestoreEdits"
    ModernButton ws, "E2", "Preserve UserEdits", "Button_RefreshDashboard_PreserveUserEdits"

    If Err.Number <> 0 Then LogUserEditsOperation "SetupDashboardUI_EndRefresh: Note - Error setting UI elements.": Err.Clear
    On Error GoTo 0
End Sub


'================================================================================
'              4. UserEdits SHEET MANAGEMENT (Save, Setup, Backup)
'              *** Using Reviewed/Robust Versions ***
'================================================================================

'------------------------------------------------------------------------------
' SaveUserEditsFromDashboard - Captures L-N from Dashboard -> UserEdits A-F
' Uses dictionary lookup and cleaned keys for reliable matching.
' *** Using robust version ***
'------------------------------------------------------------------------------
Public Sub SaveUserEditsFromDashboard()
    Dim wsDash As Worksheet, wsEdits As Worksheet
    Dim lastRowDash As Long, lastRowEdits As Long, destRow As Long
    Dim i As Long
    Dim dashDocNum As String, cleanedDashDocNum As String
    Dim userEditsDict As Object ' Dictionary for existing UserEdits lookup
    Dim editRow As Variant      ' Stores existing row number from dictionary
    Dim hasEdits As Boolean, wasChanged As Boolean
    Dim editsSavedCount As Long, editsUpdatedCount As Long

    LogUserEditsOperation "SaveUserEditsFromDashboard: Starting process..."

    ' --- Get Worksheet Objects ---
    Set wsDash = GetOrCreateDashboardSheet(DASHBOARD_SHEET_NAME) ' Use helper
    If wsDash Is Nothing Then
        LogUserEditsOperation "SaveUserEditsFromDashboard: ERROR - Dashboard sheet '" & DASHBOARD_SHEET_NAME & "' not found."
        Exit Sub
    End If
    SetupUserEditsSheet ' Ensure UserEdits sheet exists and is structured correctly
    Set wsEdits = ThisWorkbook.Sheets(USEREDITS_SHEET_NAME)
    If wsEdits Is Nothing Then
        LogUserEditsOperation "SaveUserEditsFromDashboard: ERROR - UserEdits sheet '" & USEREDITS_SHEET_NAME & "' could not be accessed."
        Exit Sub
    End If

    ' --- Load Existing UserEdits for Lookup ---
    Set userEditsDict = LoadUserEditsToDictionary(wsEdits) ' Uses CleanDocumentNumber for keys
    lastRowDash = wsDash.Cells(wsDash.rows.Count, "A").End(xlUp).Row

    LogUserEditsOperation "SaveUserEditsFromDashboard: Checking dashboard rows 4 to " & lastRowDash & " for edits..."

    ' --- Loop Through Dashboard Rows ---
    If lastRowDash >= 4 Then
        For i = 4 To lastRowDash
            dashDocNum = CStr(wsDash.Cells(i, "A").Value)         ' Get raw DocNum from Dashboard
            cleanedDashDocNum = CleanDocumentNumber(dashDocNum) ' Clean it for lookup
            wasChanged = False ' Reset change tracker for each row

            If cleanedDashDocNum <> "" Then ' Only process if there's a valid DocNum

                ' Check if dashboard row L, M, or N has any data
                hasEdits = False
                If wsDash.Cells(i, DB_COL_PHASE).Value <> "" Or _
                   wsDash.Cells(i, DB_COL_LASTCONTACT).Value <> "" Or _
                   wsDash.Cells(i, DB_COL_COMMENTS).Value <> "" Then
                    hasEdits = True
                End If

                ' Find if this cleaned DocNum already exists in UserEdits
                editRow = 0 ' Reset flag
                If userEditsDict.Exists(cleanedDashDocNum) Then
                    editRow = userEditsDict(cleanedDashDocNum) ' Get existing SHEET row number
                End If

                ' --- Process If Dashboard Has Edits OR If Entry Exists in UserEdits ---
                If hasEdits Or editRow > 0 Then

                    ' Determine destination row in UserEdits sheet
                    If editRow > 0 Then
                        destRow = editRow ' Update existing row
                    Else
                        ' Add new row
                        lastRowEdits = wsEdits.Cells(wsEdits.rows.Count, UE_COL_DOCNUM).End(xlUp).Row + 1
                        If lastRowEdits < 2 Then lastRowEdits = 2 ' Ensure starting at row 2
                        destRow = lastRowEdits
                        wsEdits.Cells(destRow, UE_COL_DOCNUM).Value = cleanedDashDocNum ' Write the CLEANED document number
                        userEditsDict.Add cleanedDashDocNum, destRow ' Add to dictionary immediately
                        wasChanged = True ' New row always counts as changed
                        editsSavedCount = editsSavedCount + 1
                    End If

                    ' Get current values from Dashboard L, M, N
                    Dim dbPhaseVal, dbLastContactVal, dbCommentsVal
                    dbPhaseVal = wsDash.Cells(i, DB_COL_PHASE).Value
                    dbLastContactVal = wsDash.Cells(i, DB_COL_LASTCONTACT).Value
                    dbCommentsVal = wsDash.Cells(i, DB_COL_COMMENTS).Value

                    ' Compare with UserEdits values and update if different (use CStr for safe comparison)
                    If editRow = 0 Or CStr(wsEdits.Cells(destRow, UE_COL_PHASE).Value) <> CStr(dbPhaseVal) Then
                        wsEdits.Cells(destRow, UE_COL_PHASE).Value = dbPhaseVal
                        If editRow > 0 Then wasChanged = True
                    End If
                    If editRow = 0 Or CStr(wsEdits.Cells(destRow, UE_COL_LASTCONTACT).Value) <> CStr(dbLastContactVal) Then
                         wsEdits.Cells(destRow, UE_COL_LASTCONTACT).Value = dbLastContactVal
                         If editRow > 0 Then wasChanged = True
                    End If
                    If editRow = 0 Or CStr(wsEdits.Cells(destRow, UE_COL_COMMENTS).Value) <> CStr(dbCommentsVal) Then
                         wsEdits.Cells(destRow, UE_COL_COMMENTS).Value = dbCommentsVal
                         If editRow > 0 Then wasChanged = True
                    End If

                    ' If any change was made, update Source/Timestamp
                    If wasChanged Then
                        wsEdits.Cells(destRow, UE_COL_SOURCE).Value = Module_Identity.GetWorkbookIdentity()
                        wsEdits.Cells(destRow, UE_COL_TIMESTAMP).Value = Format$(Now(), "yyyy-mm-dd hh:nn:ss")
                        If editRow > 0 Then editsUpdatedCount = editsUpdatedCount + 1
                    End If
                End If ' End If hasEdits Or editRow > 0
            End If ' End If cleanedDashDocNum <> ""
        Next i
    End If ' End If lastRowDash >= 4

    LogUserEditsOperation "SaveUserEditsFromDashboard: Finished. New Edits Saved: " & editsSavedCount & ". Existing Edits Updated: " & editsUpdatedCount & "."
    Set userEditsDict = Nothing ' Clean up dictionary
    Set wsDash = Nothing: Set wsEdits = Nothing
    ' No Exit Sub here, allow RefreshDashboard's error handler to manage flow
    ' Exit Sub

'ErrorHandler_SaveUserEdits: ' Now handled by RefreshDashboard's main handler
'    LogUserEditsOperation "SaveUserEditsFromDashboard: ERROR [" & Err.Number & "] " & Err.Description
'    Set userEditsDict = Nothing: Set wsDash = Nothing: Set wsEdits = Nothing
End Sub


'------------------------------------------------------------------------------
' SetupUserEditsSheet - Creates or Verifies the UserEdits sheet structure (A-F)
' Includes safety checks and backup before modifying structure.
' *** Using robust version ***
'------------------------------------------------------------------------------
Public Sub SetupUserEditsSheet()
    Dim wsEdits As Worksheet
    Dim needsBackup As Boolean
    Dim wsBackup As Worksheet
    Dim currentHeaders As Variant
    Dim expectedHeaders As Variant
    Dim structureCorrect As Boolean
    Dim backupSuccess As Boolean
    Dim i As Long
    Dim emergencyFlag As Boolean

    expectedHeaders = Array("DocNumber", "Engagement Phase", "Last Contact Date", _
                            "User Comments", "ChangeSource", "Timestamp") ' A-F

    LogUserEditsOperation "SetupUserEditsSheet: Checking structure..."

    ' --- Check if Sheet Exists ---
    On Error Resume Next
    Set wsEdits = ThisWorkbook.Sheets(USEREDITS_SHEET_NAME)
    If Err.Number <> 0 Then Err.Clear ' Ignore error if sheet doesn't exist
    On Error GoTo ErrorHandler_SetupUserEdits ' Use specific handler for this sub

    If wsEdits Is Nothing Then
        ' --- Create New Sheet ---
        LogUserEditsOperation "SetupUserEditsSheet: Sheet doesn't exist. Creating new."
        Set wsEdits = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        wsEdits.Name = USEREDITS_SHEET_NAME
        With wsEdits.Range(UE_COL_DOCNUM & "1:" & UE_COL_TIMESTAMP & "1") ' A1:F1
            .Value = expectedHeaders
            .Font.Bold = True
            .Interior.Color = RGB(16, 107, 193)
            .Font.Color = RGB(255, 255, 255)
        End With
        wsEdits.Visible = xlSheetHidden
        LogUserEditsOperation "SetupUserEditsSheet: Created new sheet with standard headers."
        Exit Sub ' Done
    End If

    ' --- Sheet Exists: Verify Header Structure (A-F) ---
    structureCorrect = False
    On Error Resume Next ' Handle error reading headers (e.g., protected sheet)
    currentHeaders = wsEdits.Range(UE_COL_DOCNUM & "1:" & UE_COL_TIMESTAMP & "1").Value ' A1:F1
    If Err.Number <> 0 Then
        LogUserEditsOperation "SetupUserEditsSheet: WARNING - Error reading headers from existing sheet: " & Err.Description
        needsBackup = True ' Assume structure is bad if headers can't be read
        Err.Clear
    Else
        If IsArray(currentHeaders) Then
             If UBound(currentHeaders, 2) = UBound(expectedHeaders) + 1 Then ' Check if it has 6 columns
                 structureCorrect = True ' Assume correct until mismatch found
                 For i = LBound(expectedHeaders) To UBound(expectedHeaders)
                     If CStr(currentHeaders(1, i + 1)) <> expectedHeaders(i) Then
                         LogUserEditsOperation "SetupUserEditsSheet: Header mismatch found at column " & i + 1 & ". Expected: '" & expectedHeaders(i) & "', Found: '" & CStr(currentHeaders(1, i + 1)) & "'."
                         structureCorrect = False
                         Exit For
                     End If
                 Next i
             Else
                 LogUserEditsOperation "SetupUserEditsSheet: Incorrect number of header columns found (" & UBound(currentHeaders, 2) & "). Expected 6."
                 structureCorrect = False ' Mark as incorrect if column count wrong
             End If
        Else ' currentHeaders is not an array (e.g., single cell A1 is empty or sheet is weird)
             LogUserEditsOperation "SetupUserEditsSheet: Could not read headers as an array. Assuming structure is incorrect."
             structureCorrect = False
        End If
        needsBackup = Not structureCorrect
    End If
    On Error GoTo ErrorHandler_SetupUserEdits

    If structureCorrect Then
        LogUserEditsOperation "SetupUserEditsSheet: Structure verified as correct."
        Exit Sub ' Structure is fine, nothing more to do
    End If

    ' --- Structure Incorrect: Backup and Rebuild ---
    LogUserEditsOperation "SetupUserEditsSheet: Structure incorrect or unreadable. Attempting backup and rebuild."
    backupSuccess = CreateUserEditsBackup("StructureUpdate_" & Format(Now, "yyyymmdd_hhmmss"))

    If Not backupSuccess Then
        LogUserEditsOperation "SetupUserEditsSheet: CRITICAL - Backup failed. Aborting structure update to prevent data loss."
        MsgBox "WARNING: Could not create a backup of the '" & USEREDITS_SHEET_NAME & "' sheet." & vbCrLf & vbCrLf & _
               "The sheet structure needs updating, but no changes will be made to avoid potential data loss." & vbCrLf & _
               "Please check sheet protection or other issues preventing backup.", vbCritical, "UserEdits Structure Update Failed"
        Exit Sub ' Abort if backup failed
    End If

    LogUserEditsOperation "SetupUserEditsSheet: Backup successful. Clearing and rebuilding sheet '" & wsEdits.Name & "'."

    ' --- Clear and Recreate Headers ---
    On Error Resume Next ' Handle error during clear (e.g., protection)
    wsEdits.Cells.Clear
    If Err.Number <> 0 Then
        LogUserEditsOperation "SetupUserEditsSheet: ERROR clearing sheet after backup. Manual intervention may be required. Error: " & Err.Description
        MsgBox "ERROR: Could not clear the '" & USEREDITS_SHEET_NAME & "' sheet after creating a backup." & vbCrLf & _
               "Please check sheet protection. Manual cleanup might be needed.", vbExclamation, "Clear Error"
        Exit Sub ' Stop if clear fails
    End If
    On Error GoTo ErrorHandler_SetupUserEdits ' Restore handler

    With wsEdits.Range(UE_COL_DOCNUM & "1:" & UE_COL_TIMESTAMP & "1") ' A1:F1
        .Value = expectedHeaders
        .Font.Bold = True
        .Interior.Color = RGB(16, 107, 193)
        .Font.Color = RGB(255, 255, 255)
    End With

    LogUserEditsOperation "SetupUserEditsSheet: Rebuilt headers. Attempting to restore data from most recent backup..."

    ' --- Attempt to Restore from Backup ---
    If RestoreUserEditsFromBackup() Then
        LogUserEditsOperation "SetupUserEditsSheet: Data restored from backup after structure update."
    Else
        LogUserEditsOperation "SetupUserEditsSheet: WARNING - Could not automatically restore data from backup after structure update. Manual restore may be needed from backup sheet."
    End If

    LogUserEditsOperation "SetupUserEditsSheet: Structure update process complete."
    Exit Sub

ErrorHandler_SetupUserEdits:
    LogUserEditsOperation "SetupUserEditsSheet: CRITICAL ERROR [" & Err.Number & "] " & Err.Description & " (Line: " & Erl & ")"
    MsgBox "An critical error occurred during UserEdits Sheet Setup: " & Err.Description & vbCrLf & _
           "Please check the '" & USEREDITSLOG_SHEET_NAME & "' sheet for details.", vbCritical, "UserEdits Setup Error"
End Sub

'------------------------------------------------------------------------------
' CreateUserEditsBackup - Creates a timestamped backup copy of UserEdits
' *** Using robust version ***
'------------------------------------------------------------------------------
Public Function CreateUserEditsBackup(Optional backupSuffix As String = "") As Boolean
    Dim wsEdits As Worksheet, wsBackup As Worksheet
    Dim backupName As String
    Dim backupTimestamp As String: backupTimestamp = Format(Now, "yyyymmdd_hhmmss")

    On Error GoTo BackupErrorHandler

    ' --- Determine Backup Name ---
    If backupSuffix = "" Then
        backupName = USEREDITS_SHEET_NAME & "_Backup_" & backupTimestamp
    Else
        ' Ensure suffix doesn't create overly long name (Excel limit is 31 chars)
        If Len(USEREDITS_SHEET_NAME & "_Backup_" & backupSuffix) > 31 Then
            backupName = Left$(USEREDITS_SHEET_NAME & "_Backup_" & backupSuffix, 31)
        Else
            backupName = USEREDITS_SHEET_NAME & "_Backup_" & backupSuffix
        End If
    End If

    ' --- Check Source Sheet ---
    On Error Resume Next
    Set wsEdits = ThisWorkbook.Sheets(USEREDITS_SHEET_NAME)
    If Err.Number <> 0 Then
        LogUserEditsOperation "CreateUserEditsBackup: Source sheet '" & USEREDITS_SHEET_NAME & "' not found. Cannot create backup."
        CreateUserEditsBackup = False
        Exit Function
    End If
    On Error GoTo BackupErrorHandler ' Restore handler

    ' --- Handle Existing/New Backup Sheet ---
    Application.DisplayAlerts = False ' Suppress sheet delete prompt if overwriting
    On Error Resume Next
    Set wsBackup = ThisWorkbook.Sheets(backupName)
    If Err.Number = 0 Then ' Backup with this exact name already exists
         wsBackup.Delete ' Delete existing sheet cleanly
         If Err.Number <> 0 Then GoTo BackupFailed ' Error deleting existing sheet
         LogUserEditsOperation "CreateUserEditsBackup: Deleted existing backup sheet named '" & backupName & "'."
    End If
    Err.Clear ' Clear potential error from checking sheet existence

    ' Create new backup sheet
    Set wsBackup = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
    If wsBackup Is Nothing Then GoTo BackupFailed ' Could not add sheet
    wsBackup.Name = backupName
    If Err.Number <> 0 Then ' Failed to name the sheet (e.g., invalid chars, length)
         LogUserEditsOperation "CreateUserEditsBackup: ERROR - Failed to name backup sheet '" & backupName & "'. Using default name '" & wsBackup.Name & "'."
         backupName = wsBackup.Name ' Use the default name assigned by Excel
         Err.Clear
    End If
    Application.DisplayAlerts = True ' Restore alerts
    On Error GoTo BackupErrorHandler ' Restore handler

    ' --- Copy Data ---
    wsEdits.UsedRange.Copy wsBackup.Range("A1")
    If Err.Number <> 0 Then GoTo BackupFailed ' Error during copy

    ' --- Hide Backup ---
    wsBackup.Visible = xlSheetHidden
    LogUserEditsOperation "CreateUserEditsBackup: Successfully created backup: '" & backupName & "'."
    CreateUserEditsBackup = True
    Set wsEdits = Nothing: Set wsBackup = Nothing
    Exit Function

BackupFailed:
    LogUserEditsOperation "CreateUserEditsBackup: ERROR creating backup '" & backupName & "'. Error: " & Err.Description
    ' Attempt to delete partially created backup sheet if it exists
    If Not wsBackup Is Nothing Then
        Application.DisplayAlerts = False
        On Error Resume Next ' Ignore error if sheet cannot be deleted
        wsBackup.Delete
        On Error GoTo 0
        Application.DisplayAlerts = True
    End If
BackupErrorHandler:
    If Err.Number <> 0 And Err.Number <> 9 Then ' Log error if not already logged by BackupFailed (ignore Subscript out of range if sheet check fails)
         LogUserEditsOperation "CreateUserEditsBackup: ERROR [" & Err.Number & "] " & Err.Description
    End If
    Application.DisplayAlerts = True ' Ensure alerts are back on
    CreateUserEditsBackup = False
    Set wsEdits = Nothing: Set wsBackup = Nothing
End Function


'------------------------------------------------------------------------------
' RestoreUserEditsFromBackup - Restores UserEdits from the MOST RECENT backup
' Used primarily by the ErrorHandler in RefreshDashboard.
' *** Using robust version ***
'------------------------------------------------------------------------------
Public Function RestoreUserEditsFromBackup(Optional specificBackupName As String = "") As Boolean
    Dim wsEdits As Worksheet
    Dim wsBackup As Worksheet, candidateSheet As Worksheet
    Dim backupNameToRestore As String
    Dim latestBackupTimestamp As Double: latestBackupTimestamp = 0 ' Use timestamp for more precision
    Dim datePart As String, timePart As String, backupTimestampCurrent As Double

    On Error GoTo RestoreErrorHandler

    ' --- Find the Backup Sheet to Restore ---
    If specificBackupName <> "" Then
        backupNameToRestore = specificBackupName ' Use specified name
    Else
        ' Find the most recent backup based on timestamp in name yyyymmdd_hhmmss
        For Each candidateSheet In ThisWorkbook.Sheets
            If candidateSheet.Name Like USEREDITS_SHEET_NAME & "_Backup_????????_??????*" Then ' Match pattern yyyymmdd_hhmmss
                On Error Resume Next ' Ignore parse errors
                backupTimestampCurrent = 0
                datePart = Mid$(candidateSheet.Name, Len(USEREDITS_SHEET_NAME & "_Backup_") + 1, 8) ' yyyymmdd
                timePart = Mid$(candidateSheet.Name, Len(USEREDITS_SHEET_NAME & "_Backup_") + 10, 6) ' hhmmss
                ' Combine date and time into a sortable numeric value (YYYYMMDDHHMMSS)
                backupTimestampCurrent = Val(datePart & timePart)
                If Err.Number = 0 And backupTimestampCurrent > 0 Then ' Successfully parsed timestamp
                    If backupTimestampCurrent > latestBackupTimestamp Then
                         latestBackupTimestamp = backupTimestampCurrent
                         backupNameToRestore = candidateSheet.Name
                    End If
                End If
                Err.Clear
                On Error GoTo RestoreErrorHandler ' Restore handler
            End If
        Next candidateSheet
    End If

    If backupNameToRestore = "" Then
        LogUserEditsOperation "RestoreUserEditsFromBackup: No suitable backup sheet found to restore from."
        RestoreUserEditsFromBackup = False
        Exit Function
    End If

    ' --- Get Backup and Target Sheets ---
    On Error Resume Next
    Set wsBackup = ThisWorkbook.Sheets(backupNameToRestore)
    If wsBackup Is Nothing Then Err.Raise 9, , "Backup sheet '" & backupNameToRestore & "' not found." ' Raise error if specific sheet missing
    On Error GoTo RestoreErrorHandler

    ' Ensure UserEdits sheet exists (might have been deleted during failed process)
    SetupUserEditsSheet ' This will create it if missing
    Set wsEdits = ThisWorkbook.Sheets(USEREDITS_SHEET_NAME)
    If wsEdits Is Nothing Then Err.Raise 9, , "Target sheet '" & USEREDITS_SHEET_NAME & "' could not be accessed."

    ' --- Perform Restore ---
    LogUserEditsOperation "RestoreUserEditsFromBackup: Restoring data from '" & backupNameToRestore & "' to '" & wsEdits.Name & "'."
    wsEdits.Cells.Clear ' Clear target sheet first
    wsBackup.UsedRange.Copy wsEdits.Range("A1")

    LogUserEditsOperation "RestoreUserEditsFromBackup: Restore complete."
    RestoreUserEditsFromBackup = True
    Set wsEdits = Nothing: Set wsBackup = Nothing: Set candidateSheet = Nothing
    Exit Function

RestoreErrorHandler:
    LogUserEditsOperation "RestoreUserEditsFromBackup: ERROR [" & Err.Number & "] " & Err.Description
    RestoreUserEditsFromBackup = False
    Set wsEdits = Nothing: Set wsBackup = Nothing: Set candidateSheet = Nothing
End Function

'------------------------------------------------------------------------------
' CleanupOldBackups - Deletes backup sheets older than X days
' *** Using robust version ***
'------------------------------------------------------------------------------
Private Sub CleanupOldBackups()
    Const DAYS_TO_KEEP As Long = 7 ' Keep backups for 7 days
    Dim cutoffDate As Date: cutoffDate = Date - DAYS_TO_KEEP
    Dim backupBaseName As String: backupBaseName = USEREDITS_SHEET_NAME & "_Backup_"
    Dim oldSheets As New Collection ' Use Collection to store sheets for deletion
    Dim sh As Worksheet
    Dim datePart As String, backupDate As Date, deleteCount As Long

    LogUserEditsOperation "CleanupOldBackups: Checking for backups older than " & Format(cutoffDate, "yyyy-mm-dd") & "..."

    On Error Resume Next ' Ignore errors iterating sheets

    For Each sh In ThisWorkbook.Sheets
        If sh.Visible = xlSheetHidden And sh.Name Like backupBaseName & "????????_??????*" Then ' Match yyyymmdd_hhmmss pattern
            datePart = Mid$(sh.Name, Len(backupBaseName) + 1, 8) ' yyyymmdd
            backupDate = DateSerial(1900, 1, 1) ' Default if parse fails
            Err.Clear
            backupDate = CDate(Format(datePart, "@@@@-@@-@@")) ' Parse only date part

             If Err.Number = 0 Then ' Successfully parsed date
                 If backupDate < cutoffDate Then
                     oldSheets.Add sh ' Add sheet object to collection
                 End If
             Else
                  Err.Clear
             End If
        End If
    Next sh

    If oldSheets.Count > 0 Then
        Application.DisplayAlerts = False ' Suppress delete confirmation prompts
        For Each sh In oldSheets
            On Error Resume Next ' Ignore error deleting single sheet
            sh.Delete
            If Err.Number = 0 Then deleteCount = deleteCount + 1 Else Err.Clear
        Next sh
        Application.DisplayAlerts = True
        LogUserEditsOperation "CleanupOldBackups: Deleted " & deleteCount & " old backup sheets (older than " & DAYS_TO_KEEP & " days)."
    Else
         LogUserEditsOperation "CleanupOldBackups: No old backup sheets found for deletion."
    End If

    On Error GoTo 0 ' Restore default error handling
    Set oldSheets = Nothing: Set sh = Nothing
End Sub

'================================================================================
'              5. FORMATTING & SORTING ROUTINES
'              *** Using User's Preferred Versions ***
'================================================================================

'------------------------------------------------------------------------------
' SortDashboardData - Sorts A3:N<lastRow> by F (asc), then D (desc)
' *** Uses user's provided version ***
'------------------------------------------------------------------------------
Private Sub SortDashboardData(ws As Worksheet, lastRow As Long)
    If ws Is Nothing Or lastRow < 5 Then Exit Sub ' Need header + at least 2 data rows

    LogUserEditsOperation "SortDashboardData: Sorting range A3:" & DB_COL_COMMENTS & lastRow & "..." ' Added Log
    Dim sortRange As Range
    Set sortRange = ws.Range("A3:" & DB_COL_COMMENTS & lastRow) ' A3:N<lastRow>

    With ws.Sort
        .SortFields.Clear
        ' Key 1: First Date Pulled (Column F), Ascending
        .SortFields.Add key:=sortRange.Columns(6), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
        ' Key 2: Document Amount (Column D), Descending
        .SortFields.Add key:=sortRange.Columns(4), SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:=xlSortNormal

        .SetRange sortRange
        .Header = xlYes ' Row 3 contains headers
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        On Error Resume Next ' Handle error if sort fails (e.g., protected sheet, merged cells)
        .Apply
        If Err.Number <> 0 Then LogUserEditsOperation "SortDashboardData: ERROR applying sort. Error: " & Err.Description: Err.Clear
        On Error GoTo 0 ' Restore default error handling
    End With
    Set sortRange = Nothing
    LogUserEditsOperation "SortDashboardData: Sort applied." ' Added Log
End Sub

'------------------------------------------------------------------------------
' ApplyColorFormatting - Applies conditional formatting to Col L (Phase)
' *** Using user's provided version ***
'------------------------------------------------------------------------------
Public Sub ApplyColorFormatting(ws As Worksheet, Optional startDataRow As Long = 4)
    On Error GoTo ErrorHandler_ColorFormat ' Use specific handler
    If ws Is Nothing Then Exit Sub ' Added check

    Dim endRow As Long
    endRow = ws.Cells(ws.rows.Count, "A").End(xlUp).Row
    If endRow < startDataRow Then Exit Sub ' No data rows to format

    LogUserEditsOperation "ApplyColorFormatting: Applying CF to " & DB_COL_PHASE & startDataRow & ":" & DB_COL_PHASE & endRow & "." ' Added Log

    Dim rngPhase As Range
    Set rngPhase = ws.Range(DB_COL_PHASE & startDataRow & ":" & DB_COL_PHASE & endRow) ' Column L data range

    ' No need to unprotect/reprotect here if called correctly from RefreshDashboard
    ApplyStageFormatting rngPhase ' Call helper sub with specific rules

    Set rngPhase = Nothing
    Exit Sub

ErrorHandler_ColorFormat:
    LogUserEditsOperation "Error in ApplyColorFormatting: " & Err.Description
    Set rngPhase = Nothing
End Sub

'------------------------------------------------------------------------------
' ApplyStageFormatting - Helper containing specific conditional format rules for Phase (L)
' *** Uses user's provided version ***
'------------------------------------------------------------------------------
Private Sub ApplyStageFormatting(targetRng As Range)
    If targetRng Is Nothing Then Exit Sub
    If targetRng.Cells.CountLarge = 0 Then Exit Sub

    On Error Resume Next ' Handle errors applying formats
    targetRng.FormatConditions.Delete ' Clear existing rules first
    If Err.Number <> 0 Then LogUserEditsOperation "ApplyStageFormatting: Warning - Could not delete existing format conditions.": Err.Clear

    Dim fc As FormatCondition
    Dim formulaBase As String
    Dim firstCellAddress As String

    ' Use address of the first cell in the range (e.g., L4) for relative formulas
    firstCellAddress = targetRng.Cells(1).Address(RowAbsolute:=False, ColumnAbsolute:=False) ' Use relative row/col

    ' Using xlExpression with case-sensitive EXACT match as in user's code
    formulaBase = "=EXACT(" & firstCellAddress & ",""{PHASE}"")"

    ' --- Define Rules (Copied from user's code) ---
    Set fc = targetRng.FormatConditions.Add(Type:=xlExpression, Formula1:=Replace(formulaBase, "{PHASE}", "First F/U"))
    If Not fc Is Nothing Then fc.Interior.Color = RGB(208, 230, 245)

    Set fc = targetRng.FormatConditions.Add(Type:=xlExpression, Formula1:=Replace(formulaBase, "{PHASE}", "Second F/U"))
    If Not fc Is Nothing Then fc.Interior.Color = RGB(146, 198, 237)

    Set fc = targetRng.FormatConditions.Add(Type:=xlExpression, Formula1:=Replace(formulaBase, "{PHASE}", "Third F/U"))
    If Not fc Is Nothing Then fc.Interior.Color = RGB(245, 225, 113)

    Set fc = targetRng.FormatConditions.Add(Type:=xlExpression, Formula1:=Replace(formulaBase, "{PHASE}", "Long-Term F/U"))
    If Not fc Is Nothing Then fc.Interior.Color = RGB(255, 150, 54)

    Set fc = targetRng.FormatConditions.Add(Type:=xlExpression, Formula1:=Replace(formulaBase, "{PHASE}", "Requoting"))
    If Not fc Is Nothing Then fc.Interior.Color = RGB(227, 215, 232)

    Set fc = targetRng.FormatConditions.Add(Type:=xlExpression, Formula1:=Replace(formulaBase, "{PHASE}", "Pending"))
    If Not fc Is Nothing Then fc.Interior.Color = RGB(255, 247, 209)

    Set fc = targetRng.FormatConditions.Add(Type:=xlExpression, Formula1:=Replace(formulaBase, "{PHASE}", "No Response"))
    If Not fc Is Nothing Then fc.Interior.Color = RGB(245, 238, 224)

    Set fc = targetRng.FormatConditions.Add(Type:=xlExpression, Formula1:=Replace(formulaBase, "{PHASE}", "Texas (No F/U)"))
    If Not fc Is Nothing Then fc.Interior.Color = RGB(230, 217, 204)

    Set fc = targetRng.FormatConditions.Add(Type:=xlExpression, Formula1:=Replace(formulaBase, "{PHASE}", "AF"))
    If Not fc Is Nothing Then fc.Interior.Color = RGB(162, 217, 210)

    Set fc = targetRng.FormatConditions.Add(Type:=xlExpression, Formula1:=Replace(formulaBase, "{PHASE}", "RZ"))
    If Not fc Is Nothing Then fc.Interior.Color = RGB(138, 155, 212)

    Set fc = targetRng.FormatConditions.Add(Type:=xlExpression, Formula1:=Replace(formulaBase, "{PHASE}", "KMH"))
    If Not fc Is Nothing Then fc.Interior.Color = RGB(247, 196, 175)

    Set fc = targetRng.FormatConditions.Add(Type:=xlExpression, Formula1:=Replace(formulaBase, "{PHASE}", "RI"))
    If Not fc Is Nothing Then fc.Interior.Color = RGB(191, 225, 243)

    Set fc = targetRng.FormatConditions.Add(Type:=xlExpression, Formula1:=Replace(formulaBase, "{PHASE}", "WW/OM"))
    If Not fc Is Nothing Then fc.Interior.Color = RGB(155, 124, 185)

    Set fc = targetRng.FormatConditions.Add(Type:=xlExpression, Formula1:=Replace(formulaBase, "{PHASE}", "Converted"))
    If Not fc Is Nothing Then With fc: .Interior.Color = RGB(120, 235, 120): .Font.Bold = True: End With

    Set fc = targetRng.FormatConditions.Add(Type:=xlExpression, Formula1:=Replace(formulaBase, "{PHASE}", "Declined"))
    If Not fc Is Nothing Then fc.Interior.Color = RGB(209, 47, 47)

    Set fc = targetRng.FormatConditions.Add(Type:=xlExpression, Formula1:=Replace(formulaBase, "{PHASE}", "Closed (Extra Order)"))
    If Not fc Is Nothing Then fc.Interior.Color = RGB(184, 39, 39)

    Set fc = targetRng.FormatConditions.Add(Type:=xlExpression, Formula1:=Replace(formulaBase, "{PHASE}", "Closed"))
    If Not fc Is Nothing Then fc.Interior.Color = RGB(166, 28, 28)

    If Err.Number <> 0 Then LogUserEditsOperation "ApplyStageFormatting: ERROR applying one or more format conditions. Error: " & Err.Description: Err.Clear
    On Error GoTo 0
    Set fc = Nothing
End Sub


'------------------------------------------------------------------------------
' ApplyWorkflowLocationFormatting - Applies CF to Col J (Workflow)
' *** Uses user's provided version ***
'------------------------------------------------------------------------------
Private Sub ApplyWorkflowLocationFormatting(ws As Worksheet, Optional startDataRow As Long = 4)
    On Error GoTo ErrorHandler_WorkflowFormat ' Use specific handler

    If ws Is Nothing Then Exit Sub ' Added check
    Dim endRow As Long
    endRow = ws.Cells(ws.rows.Count, "A").End(xlUp).Row
    If endRow < startDataRow Then Exit Sub ' No data rows to format

    LogUserEditsOperation "ApplyWorkflowLocationFormatting: Applying CF to " & DB_COL_WORKFLOW_LOCATION & startDataRow & ":" & DB_COL_WORKFLOW_LOCATION & endRow & "." ' Added Log

    Dim rngLocation As Range
    Set rngLocation = ws.Range(DB_COL_WORKFLOW_LOCATION & startDataRow & ":" & DB_COL_WORKFLOW_LOCATION & endRow) ' Column J data range

    ' No need to unprotect/reprotect here if called correctly from RefreshDashboard
    On Error Resume Next ' Handle errors applying formats
    rngLocation.FormatConditions.Delete ' Clear existing rules first
    If Err.Number <> 0 Then LogUserEditsOperation "ApplyWorkflowLocationFormatting: Warning - Could not delete existing conditions.": Err.Clear

    Dim fc As FormatCondition
    ' Add rules based on exact text values (Copied from user's code)
    Set fc = rngLocation.FormatConditions.Add(Type:=xlCellValue, Operator:=xlEqual, Formula1:="=""Quote Only""")
    If Not fc Is Nothing Then fc.Interior.Color = RGB(230, 240, 248)

    Set fc = rngLocation.FormatConditions.Add(Type:=xlCellValue, Operator:=xlEqual, Formula1:="=""1. New Orders""")
    If Not fc Is Nothing Then fc.Interior.Color = RGB(208, 230, 245)

    Set fc = rngLocation.FormatConditions.Add(Type:=xlCellValue, Operator:=xlEqual, Formula1:="=""2. Open Orders""")
    If Not fc Is Nothing Then fc.Interior.Color = RGB(255, 247, 209)

    Set fc = rngLocation.FormatConditions.Add(Type:=xlCellValue, Operator:=xlEqual, Formula1:="=""3. As Available Orders""")
    If Not fc Is Nothing Then fc.Interior.Color = RGB(227, 215, 232)

    Set fc = rngLocation.FormatConditions.Add(Type:=xlCellValue, Operator:=xlEqual, Formula1:="=""4. Hold Orders""")
    If Not fc Is Nothing Then fc.Interior.Color = RGB(255, 150, 54)

    Set fc = rngLocation.FormatConditions.Add(Type:=xlCellValue, Operator:=xlEqual, Formula1:="=""5. Closed Files""")
    If Not fc Is Nothing Then fc.Interior.Color = RGB(230, 230, 230)

    Set fc = rngLocation.FormatConditions.Add(Type:=xlCellValue, Operator:=xlEqual, Formula1:="=""6. Declined Orders""")
    If Not fc Is Nothing Then fc.Interior.Color = RGB(209, 47, 47)

    If Err.Number <> 0 Then LogUserEditsOperation "ApplyWorkflowLocationFormatting: ERROR applying conditions. Error: " & Err.Description: Err.Clear
    On Error GoTo 0 ' Restore default handler

    Set rngLocation = Nothing: Set fc = Nothing
    Exit Sub

ErrorHandler_WorkflowFormat:
    LogUserEditsOperation "Error in ApplyWorkflowLocationFormatting: " & Err.Description
    Set rngLocation = Nothing: Set fc = Nothing
End Sub


'------------------------------------------------------------------------------
' ProtectUserColumns - Locks columns A-K, Unlocks L-N, Protects sheet
' *** Uses user's provided version ***
'------------------------------------------------------------------------------
Public Sub ProtectUserColumns(ws As Worksheet)
    If ws Is Nothing Then Exit Sub ' Added check
    LogUserEditsOperation "ProtectUserColumns: Locking A-K, Unlocking L-N, Protecting sheet '" & ws.Name & "'." ' Added Log
    On Error Resume Next ' Ignore errors if already protected/unprotected

    ws.Unprotect ' Ensure unprotected first

    ws.Cells.Locked = True ' Lock all

    Dim lastDataRowProtected As Long ' Use different variable name
    lastDataRowProtected = ws.Cells(ws.rows.Count, "A").End(xlUp).Row
    If lastDataRowProtected >= 4 Then
        ws.Range(DB_COL_PHASE & "4:" & DB_COL_COMMENTS & lastDataRowProtected).Locked = False ' Unlock L:N
    End If

    ' Protect sheet allowing selection of locked/unlocked cells
    ws.Protect UserInterfaceOnly:=True, DrawingObjects:=True, Contents:=True, Scenarios:=True

    If Err.Number <> 0 Then LogUserEditsOperation "ProtectUserColumns: ERROR applying protection. Error: " & Err.Description: Err.Clear
    On Error GoTo 0
End Sub


'================================================================================
'              6. TEXT-ONLY DASHBOARD COPY ROUTINE
'              *** Using Reviewed Version ***
'================================================================================

'------------------------------------------------------------------------------
' CreateOrUpdateTextOnlySheet - Creates a values-only copy of the dashboard
'------------------------------------------------------------------------------
Private Sub CreateOrUpdateTextOnlySheet(wsSource As Worksheet)
    If wsSource Is Nothing Then Exit Sub ' Exit if source worksheet is invalid
    LogUserEditsOperation "CreateOrUpdateTextOnlySheet: Starting process..."

    Dim wsValues As Worksheet, currentSheet As Worksheet
    Dim lastRowSource As Long, lastRowValues As Long
    Dim srcRange As Range

    On Error GoTo TextOnlyErrorHandler ' Enable error handling for this sub
    Set currentSheet = ActiveSheet ' Remember active sheet

    ' --- Get or Create Text-Only Sheet ---
    On Error Resume Next ' Temporarily ignore errors for sheet check/add
    Set wsValues = ThisWorkbook.Sheets(TEXT_ONLY_SHEET_NAME)
    If wsValues Is Nothing Then ' Sheet doesn't exist, create it
        Set wsValues = ThisWorkbook.Sheets.Add(After:=wsSource)
        If wsValues Is Nothing Then GoTo TextOnlyFatalError ' Could not add sheet
        wsValues.Name = TEXT_ONLY_SHEET_NAME
        If Err.Number <> 0 Then ' Naming failed (e.g., conflict)
            LogUserEditsOperation "CreateOrUpdateTextOnlySheet: Warning - Failed to name Text-Only sheet '" & TEXT_ONLY_SHEET_NAME & "'. Using default name '" & wsValues.Name & "'."
            Err.Clear
        Else
             LogUserEditsOperation "CreateOrUpdateTextOnlySheet: Created new sheet: '" & TEXT_ONLY_SHEET_NAME & "'."
        End If
    End If
    On Error GoTo TextOnlyErrorHandler ' Restore main error handling

    ' --- Ensure sheet is valid and prepare ---
    If wsValues Is Nothing Then GoTo TextOnlyFatalError ' Should not happen, but safety check
    wsValues.Visible = xlSheetVisible ' Ensure visible for operations
    wsValues.Cells.Clear ' Clear existing content and formats

    ' --- Copy Data (Values and Number Formats) ---
    lastRowSource = wsSource.Cells(wsSource.rows.Count, "A").End(xlUp).Row
    If lastRowSource >= 3 Then ' Ensure there's at least header data on source
        Set srcRange = wsSource.Range("A3:" & DB_COL_COMMENTS & lastRowSource) ' A3:N<lastRow> (Headers + Data)
        srcRange.Copy
        wsValues.Range("A1").PasteSpecial Paste:=xlPasteValuesAndNumberFormats ' Paste Headers+Data starting at A1
        wsValues.Range("A1").PasteSpecial Paste:=xlPasteColumnWidths        ' Paste Column Widths
        Application.CutCopyMode = False
        LogUserEditsOperation "CreateOrUpdateTextOnlySheet: Pasted headers & data (Rows " & 3 & "-" & lastRowSource & " from source) to '" & wsValues.Name & "'."
    Else
        LogUserEditsOperation "CreateOrUpdateTextOnlySheet: No source data found on '" & wsSource.Name & "' (A3:N) to copy. Only setting up headers."
        ' Setup headers manually if no data copied
        With wsValues.Range("A1:" & DB_COL_COMMENTS & "1") ' A1:N1
            .Value = wsSource.Range("A3:" & DB_COL_COMMENTS & "3").Value ' Copy header text
            .Font.Bold = wsSource.Range("A3").Font.Bold
            .Interior.Color = wsSource.Range("A3").Interior.Color
            .Font.Color = wsSource.Range("A3").Font.Color
            .HorizontalAlignment = xlCenter
        End With
    End If

    ' --- Re-apply Conditional Formatting (Data starts at row 2 on wsValues) ---
    lastRowValues = wsValues.Cells(wsValues.rows.Count, "A").End(xlUp).Row
    If lastRowValues >= 2 Then ' Headers are row 1, data starts row 2
        LogUserEditsOperation "CreateOrUpdateTextOnlySheet: Applying conditional formatting..."
        ' Pass the specific range to ApplyStageFormatting
        ApplyStageFormatting wsValues.Range(DB_COL_PHASE & "2:" & DB_COL_PHASE & lastRowValues)
        ' Pass sheet and START row to ApplyWorkflowLocationFormatting
        ApplyWorkflowLocationFormatting wsValues, 2 ' Start formatting from row 2
        LogUserEditsOperation "CreateOrUpdateTextOnlySheet: Applied conditional formatting to rows 2:" & lastRowValues & "."
    End If

    ' --- Final Formatting for Text-Only sheet ---
    wsValues.Columns("A:" & DB_COL_COMMENTS).VerticalAlignment = xlTop ' Align top
    wsValues.rows(1).HorizontalAlignment = xlCenter ' Center Headers

    ' --- Ensure Unprotected and Unfrozen ---
    On Error Resume Next ' Temporarily ignore errors for Unprotect/Freeze
    wsValues.Unprotect
    If ActiveSheet.Name = wsValues.Name Then ActiveWindow.FreezePanes = False
    On Error GoTo TextOnlyErrorHandler ' Restore main error handling

    ' --- Activate original sheet if needed ---
    If Not currentSheet Is Nothing Then
        If ActiveSheet.Name <> currentSheet.Name Then currentSheet.Activate
    End If
    LogUserEditsOperation "CreateOrUpdateTextOnlySheet: Finalized sheet '" & wsValues.Name & "'."

    ' --- Cleanup ---
    Set currentSheet = Nothing: Set wsValues = Nothing: Set srcRange = Nothing
    Exit Sub ' Normal exit

TextOnlyFatalError:
    LogUserEditsOperation "CreateOrUpdateTextOnlySheet: CRITICAL ERROR - Could not get or create the Text-Only sheet. Aborting Text-Only update."
    GoTo TextOnlyCleanup

TextOnlyErrorHandler:
    LogUserEditsOperation "CreateOrUpdateTextOnlySheet: ERROR [" & Err.Number & "] " & Err.Description & " (Line: " & Erl & ")"
TextOnlyCleanup: ' Label used by GoTo if sheet add failed or other fatal error
    Application.CutCopyMode = False ' Ensure copy mode is off
    On Error Resume Next ' Use Resume Next for final cleanup attempts
    If Not currentSheet Is Nothing Then
        If ActiveSheet.Name <> currentSheet.Name Then currentSheet.Activate
    End If
    Set currentSheet = Nothing: Set wsValues = Nothing: Set srcRange = Nothing
End Sub


'================================================================================
'              7. Legacy/Compatibility (Optional)
'================================================================================

' Keep this if older buttons/macros might still call the old name
Public Sub CreateSQRCTDashboard()
    LogUserEditsOperation "Legacy 'CreateSQRCTDashboard' called. Redirecting to 'RefreshDashboard_SaveAndRestoreEdits'."
    RefreshDashboard PreserveUserEdits:=False
End Sub

'------------------------------------------------------------------------------
' Module_Identity Placeholder (Ensure this module exists in your project)
'------------------------------------------------------------------------------
' If you don't have Module_Identity, create it and add this code:
'
' Option Explicit
' Public Const WORKBOOK_IDENTITY As String = "UNIQUE_ID" ' e.g., "RZ", "AF", "MASTER"
' Public Function GetWorkbookIdentity() As String
'    GetWorkbookIdentity = WORKBOOK_IDENTITY
' End Function
'------------------------------------------------------------------------------

