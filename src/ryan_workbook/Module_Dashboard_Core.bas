Option Explicit

'===============================================================================
' MODULE_DASHBOARD_CORE
' Contains the primary logic for refreshing the SQRCT Dashboard display,
' including data fetching, sorting, formatting, and UI setup.
' Calls functions in Module_Dashboard_UserEdits for persistence and logging.
'===============================================================================

' --- PUBLIC CONSTANTS (Accessible by other modules) ---
Public Const DASHBOARD_SHEET_NAME As String = "SQRCT Dashboard"
Public Const USEREDITS_SHEET_NAME As String = "UserEdits" ' Used by UserEdits Module
Public Const USEREDITSLOG_SHEET_NAME As String = "UserEditsLog" ' Used by UserEdits Module
Public Const MASTER_QUOTES_FINAL_SOURCE As String = "MasterQuotes_Final" ' Name of the PQ query/table for A-I & K

' *** Constants for Power Query Output (Workflow Location) ***
Public Const PQ_LATEST_LOCATION_SHEET As String = "DocNum_LatestLocation" ' Source Sheet Name
Public Const PQ_LATEST_LOCATION_TABLE As String = "DocNum_LatestLocation" ' Source Table Name (Assumed same as sheet)
Public Const PQ_DOCNUM_COL_NAME As String = "PrimaryDocNumber"          ' Source DocNum Column Header
Public Const PQ_LOCATION_COL_NAME As String = "MostRecent_FolderLocation" ' Source Location Column Header

' UserEdits Columns (A-F) - Referenced by UserEdits Module
Public Const UE_COL_DOCNUM As String = "A"
Public Const UE_COL_PHASE As String = "B"
Public Const UE_COL_LASTCONTACT As String = "C"
Public Const UE_COL_COMMENTS As String = "D"
Public Const UE_COL_SOURCE As String = "E"
Public Const UE_COL_TIMESTAMP As String = "F"

' Dashboard Columns (A-N, Adjusted Order)
' A-I populated by MasterQuotes_Final
Public Const DB_COL_WORKFLOW_LOCATION As String = "J" ' Populated by PQ_LATEST_LOCATION lookup
Public Const DB_COL_MISSING_QUOTE As String = "K"     ' Populated by MasterQuotes_Final (Static Text)
Public Const DB_COL_PHASE As String = "L"             ' User Editable
Public Const DB_COL_LASTCONTACT As String = "M"       ' User Editable
Public Const DB_COL_COMMENTS As String = "N"          ' User Editable
' --- End Constants ---


'===============================================================================
' BUTTON FUNCTIONS: Assign these subs to buttons on the dashboard
'===============================================================================

Public Sub Button_RefreshDashboard_SaveAndRestoreEdits()
    RefreshDashboard_TwoWaySync
End Sub

Public Sub Button_RefreshDashboard_PreserveUserEdits()
    RefreshDashboard_OneWayFromUserEdits
End Sub

'===============================================================================
' MAIN FUNCTIONS: Orchestrate the refresh process
'===============================================================================

' Standard: Save Dashboard Edits -> Refresh -> Restore UserEdits
Public Sub RefreshDashboard_TwoWaySync()
    Call RefreshDashboard(PreserveUserEdits:=False)
End Sub

' Preserve: Refresh -> Apply UserEdits (doesn't save current dashboard state first)
Public Sub RefreshDashboard_OneWayFromUserEdits()
    Call RefreshDashboard(PreserveUserEdits:=True)
End Sub


'===============================================================================
' MAIN SUB: Creates or refreshes the SQRCT Dashboard
' Incorporates all fixes including post-sort Workflow Location population.
' Calls Module_Dashboard_UserEdits for relevant tasks.
'===============================================================================
Public Sub RefreshDashboard(Optional PreserveUserEdits As Boolean = False)
    Dim ws As Worksheet, wsEdits As Worksheet
    Dim lastRow As Long, lastRowEdits As Long
    Dim docNum As String
    Dim i As Long, j As Long ' Loop counters
    Dim backupCreated As Boolean
    Dim t_start As Single, t_save As Single, t_populate As Single, t_load As Single, t_restore As Single, t_format As Single, t_textOnly As Single, t_location As Single ' Timing variables
    Dim userEditsDict As Object ' Dictionary for UserEdits lookup (DocNum -> Sheet Row Number)
    Dim editSheetRow As Long ' Stores SHEET row number from dictionary
    Dim editArrayRow As Long ' Stores corresponding 1-based ARRAY row index

    ' Array variables for optimization
    Dim dashboardDocNumArray As Variant ' Used for RESTORE step
    Dim userEditsDataArray As Variant   ' Used for RESTORE step
    Dim outputEditsArray As Variant     ' Used for RESTORE step
    Dim numDashboardRows As Long        ' Used for RESTORE step

    ' Variables for Text-Only sheet
    Dim wsValues As Worksheet
    Dim srcRange As Range
    Const TEXT_ONLY_SHEET_NAME As String = "SQRCT Dashboard (Text-Only)"
    Dim currentSheet As Worksheet ' To remember active sheet

    ' Create error recovery backup before any operations
    ' *** CALLS FUNCTION IN Module_Dashboard_UserEdits ***
    backupCreated = Module_Dashboard_UserEdits.CreateUserEditsBackup("RefreshDashboard_" & Format(Now, "yyyymmdd_hhmmss"))
    Module_Dashboard_UserEdits.LogUserEditsOperation "Starting dashboard refresh. PreserveUserEdits=" & PreserveUserEdits & ", Backup created: " & backupCreated

    t_start = Timer ' Start total timer
    Debug.Print "Start Refresh: " & t_start

    On Error GoTo ErrorHandler

    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual ' Turn off calculation
    Application.DisplayAlerts = False ' Suppress alerts during operation

    ' === PREPARATION ===
    ' 1. Ensure UserEdits sheet exists with standardized structure
    ' *** CALLS FUNCTION IN Module_Dashboard_UserEdits ***
    Module_Dashboard_UserEdits.SetupUserEditsSheet
    On Error Resume Next ' Handle case where UserEdits sheet might still fail creation
    Set wsEdits = ThisWorkbook.Sheets(USEREDITS_SHEET_NAME)
    On Error GoTo ErrorHandler
    If wsEdits Is Nothing Then
        MsgBox "Critical Error: Could not find or create the '" & USEREDITS_SHEET_NAME & "' sheet. Aborting refresh.", vbCritical, "Refresh Aborted"
        GoTo Cleanup ' Cannot proceed without UserEdits
    End If

    ' 2. Save current dashboard edits to UserEdits (IF standard refresh)
    If Not PreserveUserEdits Then
        t_save = Timer
        ' *** CALLS FUNCTION IN Module_Dashboard_UserEdits ***
        Module_Dashboard_UserEdits.SaveUserEditsFromDashboard
        Debug.Print "SaveUserEdits Time: " & Timer - t_save
    End If

    ' 3. Get or Create Dashboard Sheet
    Set ws = GetOrCreateDashboardSheet(DASHBOARD_SHEET_NAME) ' Uses Private helper in this module
    If ws Is Nothing Then
         MsgBox "Critical Error: Could not find or create the '" & DASHBOARD_SHEET_NAME & "' sheet. Aborting refresh.", vbCritical, "Refresh Aborted"
         GoTo Cleanup ' Cannot proceed without Dashboard sheet
    End If

    ' === DATA POPULATION & INITIAL FORMATTING ===
    ' 4. Ensure sheet is unprotected
    On Error Resume Next ' Ignore error if already unprotected
    ws.Unprotect
    On Error GoTo ErrorHandler ' Restore error handling

    ' 5. Clean up layout issues (duplicate headers etc.)
    CleanupDashboardLayout ws ' Uses Private helper in this module

    ' 6. Initialize layout (clear data area A4:N<end>, set headers/widths)
    InitializeDashboardLayout ws ' Uses Private helper in this module

    ' 7. Populate columns A-I & K with data from MasterQuotes_Final source
    If IsMasterQuotesFinalPresent Then ' Uses Private helper in this module
        t_populate = Timer
        PopulateMasterQuotesData ws ' Uses Private helper in this module
        Debug.Print "PopulateMasterQuotes Time: " & Timer - t_populate
    Else
        MsgBox "Warning: '" & MASTER_QUOTES_FINAL_SOURCE & "' not found. Dashboard created but columns A-I and K could not be populated." & vbCrLf & _
                 "Please ensure the data source exists and is named correctly.", vbExclamation, "Data Source Not Found"
    End If

    ' 8. Determine last row with data (check Col A AFTER population)
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row

    ' 8a. Convert formulas in A-I and K to values before sorting
    If lastRow >= 4 Then
        With ws.Range("A4:" & DB_COL_COMMENTS & lastRow) ' Entire potential data range A:N
             ' Selectively convert formulas in A-I and K
             Intersect(.Cells, ws.Range("A:I, K:K")).Value = Intersect(.Cells, ws.Range("A:I, K:K")).Value
        End With
        Module_Dashboard_UserEdits.LogUserEditsOperation "Converted formulas in A-I & K to values (Rows 4:" & lastRow & ")."
    End If

    ' === SORTING & WORKFLOW LOCATION (ORDER IS CRITICAL) ===
    ' Exit if no data rows found after potentially populating Col A
    If lastRow < 4 Then
        Debug.Print "No data rows found on dashboard (lastRow=" & lastRow & ") after population. Skipping sort, location, and restore."
        GoTo SkipToFormatting ' Skip to final formatting/protection
    End If

    ' 9. Sort by First Date Pulled (F asc), then Document Amount (D desc)
    SortDashboardData ws, lastRow ' Uses Private helper in this module

    ' 10. Populate Workflow Location (Column J) AFTER sorting
    t_location = Timer
    PopulateWorkflowLocation ws, lastRow ' Uses Private helper in this module
    Debug.Print "PopulateWorkflowLocation (Post-Sort) Time: " & Timer - t_location

    ' === FINAL FORMATTING & USER EDITS RESTORE ===
    ' 11. AutoFit columns & fix column widths (Apply AFTER Col J is populated)
    With ws
        .Columns("A:I").AutoFit ' AutoFit A-I
        .Columns(DB_COL_WORKFLOW_LOCATION).ColumnWidth = 20 ' Fixed width for J (Workflow Location)
        ' REMOVED Fixed width for Column K
        ' REMOVED Fixed width for Column N

        .Columns(DB_COL_WORKFLOW_LOCATION & ":" & DB_COL_COMMENTS).AutoFit ' AutoFit J:N (Includes K, L, M, N)

        ' --- Set fixed or minimum widths AFTER AutoFit ---
        .Columns("C").ColumnWidth = 25 ' Customer Name (Set fixed width)
        .Columns(DB_COL_COMMENTS).ColumnWidth = 40 ' User Comments (N) - Set fixed width AFTER AutoFit

        ' Ensure Autofit didn't make critical columns too narrow
        If .Columns("A").ColumnWidth < 15 Then .Columns("A").ColumnWidth = 15 ' Doc Num
        If .Columns("B").ColumnWidth < 12 Then .Columns("B").ColumnWidth = 12 ' Client ID
        If .Columns("D").ColumnWidth < 15 Then .Columns("D").ColumnWidth = 15 ' Doc Amount
        If .Columns("F").ColumnWidth < 15 Then .Columns("F").ColumnWidth = 15 ' First Pulled
        ' Check AutoFit result for K and set minimum width if needed
        If .Columns(DB_COL_MISSING_QUOTE).ColumnWidth < 25 Then .Columns(DB_COL_MISSING_QUOTE).ColumnWidth = 25 ' Min width for K
        ' REMOVED minimum width check for N as fixed width is set above
    End With


    ' 12. Restore user data (L, M, N) from UserEdits to Dashboard using Arrays
    t_load = Timer
    lastRowEdits = wsEdits.Cells(wsEdits.Rows.Count, UE_COL_DOCNUM).End(xlUp).Row

    ' Read Dashboard DocNums (Column A - now SORTED)
    On Error Resume Next
    dashboardDocNumArray = ws.Range("A4:A" & lastRow).Value
    If Err.Number <> 0 Or Not IsArray(dashboardDocNumArray) Then
        Debug.Print "Error reading dashboard DocNums or no data after sort. Skipping restore."
        Err.Clear
        GoTo SkipToFormatting ' Skip restore if array read fails
    End If
    On Error GoTo ErrorHandler ' Restore error handling

    ' Read UserEdits data (Columns A to F)
    If lastRowEdits > 1 Then
        Dim userEditsRange As Range
        Set userEditsRange = wsEdits.Range(UE_COL_DOCNUM & "2:" & UE_COL_TIMESTAMP & lastRowEdits) ' A:F
        On Error Resume Next ' Handle errors during read
        If userEditsRange.Rows.Count = 1 Then ' Single data row case
            Dim singleRowData(1 To 1, 1 To 6) As Variant
            Dim cellIdx As Long
            For cellIdx = 1 To 6: singleRowData(1, cellIdx) = userEditsRange.Cells(1, cellIdx).Value: Next cellIdx
            userEditsDataArray = singleRowData
        Else ' Multiple rows
            userEditsDataArray = userEditsRange.Value
        End If
        If Err.Number <> 0 Then Module_Dashboard_UserEdits.LogUserEditsOperation "Warning: Error reading UserEdits data: " & Err.Description: Err.Clear
        On Error GoTo ErrorHandler
    Else
        Debug.Print "UserEdits sheet has no data rows (lastRowEdits=" & lastRowEdits & ")."
    End If

    ' Load UserEdits Dictionary (DocNum -> Sheet Row Number)
    ' *** CALLS FUNCTION IN Module_Dashboard_UserEdits ***
    Set userEditsDict = Module_Dashboard_UserEdits.LoadUserEditsToDictionary(wsEdits)
    Debug.Print "Load Arrays & Dictionary Time: " & Timer - t_load

    ' Initialize Output Array (for Dashboard columns L-N)
    numDashboardRows = UBound(dashboardDocNumArray, 1)
    ReDim outputEditsArray(1 To numDashboardRows, 1 To 3) ' Phase, LastContact, Comments

    ' Pre-fill output array with blanks
    For i = 1 To numDashboardRows: For j = 1 To 3: outputEditsArray(i, j) = vbNullString: Next j: Next i

    ' Process arrays to populate outputEditsArray
    t_restore = Timer
    If Not IsEmpty(userEditsDataArray) And Not IsEmpty(dashboardDocNumArray) And lastRowEdits > 1 And numDashboardRows > 0 Then
        Dim userEditsUBound As Long
        userEditsUBound = UBound(userEditsDataArray, 1)

        For i = 1 To numDashboardRows ' Loop through SORTED DASHBOARD rows (via array)
            docNum = Trim(CStr(dashboardDocNumArray(i, 1))) ' DocNum from sorted dashboard

            If docNum <> "" Then
                If userEditsDict.Exists(docNum) Then
                    editSheetRow = userEditsDict(docNum) ' Get UserEdits SHEET row number
                    editArrayRow = editSheetRow - 1      ' Calculate corresponding UserEdits ARRAY row index

                    If editArrayRow >= 1 And editArrayRow <= userEditsUBound Then
                        On Error Resume Next ' Handle potential errors during copy
                        outputEditsArray(i, 1) = userEditsDataArray(editArrayRow, 2) ' Phase (UserEdits B) -> Output Col 1
                        outputEditsArray(i, 2) = userEditsDataArray(editArrayRow, 3) ' LastContact (UserEdits C) -> Output Col 2
                        outputEditsArray(i, 3) = userEditsDataArray(editArrayRow, 4) ' Comments (UserEdits D) -> Output Col 3
                        If Err.Number <> 0 Then Debug.Print "Error copying data for DocNum '" & docNum & "'. Error: " & Err.Description: Err.Clear
                        On Error GoTo ErrorHandler
                    Else
                         Debug.Print "Warning: DocNum '" & docNum & "' UserEdits Array Row " & editArrayRow & " out of bounds (UBound=" & userEditsUBound & ")."
                    End If
                ' Else: DocNum not in UserEdits, output remains blank (already pre-filled)
                End If
            End If
        Next i
    Else
        Debug.Print "Skipping restore loop - no UserEdits data or no Dashboard data."
    End If
    Debug.Print "Restore Edits (Array Processing) Time: " & Timer - t_restore

    ' Write the output array back to the dashboard (Range L:N)
    If numDashboardRows > 0 Then
        On Error Resume Next
        ws.Range(DB_COL_PHASE & "4").Resize(numDashboardRows, 3).Value = outputEditsArray ' Write to L4:N<end>
        If Err.Number <> 0 Then Module_Dashboard_UserEdits.LogUserEditsOperation "Error writing restored edits back to dashboard: " & Err.Description: Err.Clear
        On Error GoTo ErrorHandler
    End If

    ' Clean up restore arrays and dictionary
    Set userEditsDict = Nothing
    If IsArray(dashboardDocNumArray) Then Erase dashboardDocNumArray
    If IsArray(userEditsDataArray) Then Erase userEditsDataArray
    If IsArray(outputEditsArray) Then Erase outputEditsArray

SkipToFormatting: ' Label to jump to if data population/sort/restore skipped

    ' === FINAL UI & PROTECTION ===
    ' 13. Freeze header rows
    FreezeDashboard ws ' Uses Private helper in this module

    ' 14. Apply conditional formatting & Protect columns
    t_format = Timer
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row ' Re-check last row before formatting
    If lastRow >= 4 Then ' Only apply formatting if there's data
        ApplyColorFormatting ws, lastRow           ' Applies formatting to column L (Phase)
        ApplyWorkflowLocationFormatting ws, lastRow ' Apply to Workflow Location column (J)
    End If
    ProtectUserColumns ws ' Lock A-K, Unlock L-N, Protect Sheet - Uses Private helper
    Debug.Print "Format/Protect Time: " & Timer - t_format

    ' 15. Update the timestamp
    With ws.Range("G2:I2")
        .Merge
        .Value = "Last Refreshed: " & Format$(Now(), "mm/dd/yyyy h:mm") & " MST"
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter ' Ensure vertical alignment
        .Font.Size = 9
        .Font.Name = "Segoe UI"
        .Font.Color = RGB(80, 80, 80)
    End With

    ' 16. Re-create buttons (Using simpler deletion method as workaround for compile error)
    On Error Resume Next ' Ignore errors during deletion loop
    Dim shp As Shape
    For Each shp In ws.Shapes
        ' Original simpler logic: Delete ANY shape whose top-left cell is in row 2
        If shp.TopLeftCell.Row = 2 Then
            shp.Delete
        End If
    Next shp
    On Error GoTo ErrorHandler ' Restore default error handling

    ' Re-create the buttons using ModernButton (Private helper in this module)
    ModernButton ws, "C2", "Standard Refresh", "Button_RefreshDashboard_SaveAndRestoreEdits"
    ModernButton ws, "E2", "Preserve UserEdits", "Button_RefreshDashboard_PreserveUserEdits"


    ' === TEXT-ONLY SHEET ===
    ' 17. Create/Update Text-Only Dashboard
    t_textOnly = Timer
    If Not ws Is Nothing Then CreateOrUpdateTextOnlySheet (ws) ' Call Private helper
    Debug.Print "Create Text-Only Sheet Time: " & Timer - t_textOnly

    ' === FINAL CLEANUP & MESSAGE ===
    ' 18. Final Cleanup (e.g., remove CF from Col I)
    On Error Resume Next ' Ignore errors if sheets don't exist or are protected
    ws.Unprotect
    ws.Columns("I").FormatConditions.Delete
    ws.Protect UserInterfaceOnly:=True, DrawingObjects:=True, Contents:=True, Scenarios:=True ' Re-protect main sheet

    Set wsValues = Nothing ' Reset variable
    Set wsValues = ThisWorkbook.Sheets(TEXT_ONLY_SHEET_NAME) ' Try to get ref again
    If Not wsValues Is Nothing Then
        wsValues.Unprotect ' Ensure unprotected before modifying
        wsValues.Columns("I").FormatConditions.Delete
        ' Optional: Protect Text-Only sheet
        ' wsValues.Protect UserInterfaceOnly:=True, DrawingObjects:=True, Contents:=True, Scenarios:=True
    End If
    On Error GoTo ErrorHandler
    Module_Dashboard_UserEdits.LogUserEditsOperation "Final cleanup: Ensured CF cleared from Pull Count column (I) on both sheets."

    ' 19. Build and show completion message
    Dim msgText As String
    If PreserveUserEdits Then
        msgText = DASHBOARD_SHEET_NAME & " refreshed successfully!" & vbCrLf & vbCrLf & _
                  "User edits from the '" & USEREDITS_SHEET_NAME & "' sheet were preserved and applied to the dashboard." & vbCrLf & _
                  "(No changes made on the dashboard itself were saved back to '" & USEREDITS_SHEET_NAME & "' during this refresh.)"
    Else
        msgText = DASHBOARD_SHEET_NAME & " refreshed successfully!" & vbCrLf & vbCrLf & _
                  "Any edits made directly on the dashboard were saved to '" & USEREDITS_SHEET_NAME & "' before the refresh." & vbCrLf & _
                  "All edits from '" & USEREDITS_SHEET_NAME & "' were then restored to the dashboard."
    End If
    MsgBox msgText, vbInformation, "Dashboard Refresh Complete"

    ' 20. Log successful completion
    Module_Dashboard_UserEdits.LogUserEditsOperation "Dashboard refresh completed successfully. Mode: " & IIf(PreserveUserEdits, "PreserveUserEdits", "StandardRefresh")
    Debug.Print "Total RefreshDashboard VBA Time: " & Timer - t_start

    ' 21. Clean up old backups (optional policy, e.g., older than 7 days)
    ' *** CALLS FUNCTION IN Module_Dashboard_UserEdits ***
    If backupCreated Then Module_Dashboard_UserEdits.CleanupOldBackups


Cleanup:
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Application.Calculation = xlCalculationAutomatic ' Restore calculation
    Application.DisplayAlerts = True ' Restore alerts
    ' Clean up object variables
    Set userEditsDict = Nothing
    Set wsValues = Nothing
    Set srcRange = Nothing
    Set currentSheet = Nothing
    Set ws = Nothing
    Set wsEdits = Nothing
    ' Erase arrays if necessary (less critical now as scope ends)
    If IsArray(dashboardDocNumArray) Then Erase dashboardDocNumArray
    If IsArray(userEditsDataArray) Then Erase userEditsDataArray
    If IsArray(outputEditsArray) Then Erase outputEditsArray
    Exit Sub

ErrorHandler:
    Module_Dashboard_UserEdits.LogUserEditsOperation "ERROR in RefreshDashboard: [" & Err.Number & "] " & Err.Description & " (Line: " & Erl & ")"
    MsgBox "An error occurred during dashboard refresh." & vbCrLf & vbCrLf & _
           "Error: " & Err.Description & vbCrLf & _
           "(Error Code: " & Err.Number & ")", vbCritical, "Dashboard Refresh Error"
    ' Attempt to restore from the most recent backup if available
    If backupCreated Then
        ' *** CALLS FUNCTION IN Module_Dashboard_UserEdits ***
        If Module_Dashboard_UserEdits.RestoreUserEditsFromBackup() Then ' Find most recent backup
            MsgBox "Attempted to restore your UserEdits from the most recent backup.", vbInformation, "Backup Restored"
        Else
             MsgBox "Attempted to restore UserEdits from backup, but failed. Please check for backup sheets manually.", vbExclamation, "Backup Restore Failed"
        End If
    End If
    Resume Cleanup ' Go to cleanup section after error
End Sub


'===============================================================================
' PRIVATE HELPER FUNCTIONS FOR Module_Dashboard_Core
'===============================================================================

' --- Data Population Helpers ---

Private Sub PopulateWorkflowLocation(ws As Worksheet, lastRow As Long)
    ' FINAL RECOMMENDED VERSION: Runs AFTER sorting, uses Dictionary, writes row-by-row.
    Dim wsPQ As Worksheet, tblPQ As ListObject, i As Long, docNum As String
    Dim locationResult As Variant, t_start As Single, pqDict As Object
    Dim pqRowCount As Long, sourceDataArray As Variant

    t_start = Timer
    Module_Dashboard_UserEdits.LogUserEditsOperation "Starting PopulateWorkflowLocation (Post-Sort, Row-by-Row Write)"
    On Error GoTo LocationErrorHandler

    ' Get References to PQ Output Sheet and Table
    Set wsPQ = Nothing: On Error Resume Next
    Set wsPQ = ThisWorkbook.Sheets(PQ_LATEST_LOCATION_SHEET): On Error GoTo LocationErrorHandler
    If wsPQ Is Nothing Then GoTo SourceSheetError

    Set tblPQ = Nothing: On Error Resume Next
    Set tblPQ = wsPQ.ListObjects(PQ_LATEST_LOCATION_TABLE)
    If tblPQ Is Nothing Then Set tblPQ = wsPQ.ListObjects(1) ' Fallback
    On Error GoTo LocationErrorHandler
    If tblPQ Is Nothing Then GoTo SourceTableError

    ' Verify Required Columns Exist in PQ Table
    Dim pqDocNumColIndex As Long, pqLocationColIndex As Long
    On Error Resume Next
    pqDocNumColIndex = tblPQ.ListColumns(PQ_DOCNUM_COL_NAME).Index
    pqLocationColIndex = tblPQ.ListColumns(PQ_LOCATION_COL_NAME).Index
    On Error GoTo LocationErrorHandler
    If pqDocNumColIndex = 0 Or pqLocationColIndex = 0 Then GoTo SourceColumnError

    ' Read Source Data into Dictionary
    Set pqDict = CreateObject("Scripting.Dictionary")
    pqDict.CompareMode = vbTextCompare

    On Error Resume Next ' Handle errors accessing table data
    If tblPQ.DataBodyRange Is Nothing Then pqRowCount = 0 Else
        sourceDataArray = tblPQ.ListColumns(Array(PQ_DOCNUM_COL_NAME, PQ_LOCATION_COL_NAME)).DataBodyRange.Value
        If Err.Number <> 0 Then ' Fallback to reading whole table if specific columns fail
            Err.Clear: sourceDataArray = tblPQ.Range.Value
            If Err.Number <> 0 Or Not IsArray(sourceDataArray) Then GoTo SourceReadError
            Dim r As Long, c As Long, docCol As Long, locCol As Long
            For c = 1 To UBound(sourceDataArray, 2) ' Find columns in full array
                 If sourceDataArray(1, c) = PQ_DOCNUM_COL_NAME Then docCol = c
                 If sourceDataArray(1, c) = PQ_LOCATION_COL_NAME Then locCol = c
                 If docCol > 0 And locCol > 0 Then Exit For
            Next c
            If docCol = 0 Or locCol = 0 Then GoTo SourceHeaderError
            pqRowCount = UBound(sourceDataArray, 1) - 1
            If pqRowCount > 0 Then ' Populate dict from full array
                For r = 2 To UBound(sourceDataArray, 1)
                    Dim dictKey As String: dictKey = Trim(CStr(sourceDataArray(r, docCol)))
                    If Len(dictKey) > 0 And Not pqDict.Exists(dictKey) Then pqDict.Add dictKey, sourceDataArray(r, locCol)
                Next r
            End If
        Else ' Successfully read specific columns
             If IsArray(sourceDataArray) Then
                 pqRowCount = UBound(sourceDataArray, 1)
                 Dim dictKey As String, dictValue As Variant, r As Long
                 For r = 1 To pqRowCount ' Populate dict from 2-col array
                     dictKey = Trim(CStr(sourceDataArray(r, 1)))
                     dictValue = sourceDataArray(r, 2)
                     If Len(dictKey) > 0 And Not pqDict.Exists(dictKey) Then pqDict.Add dictKey, dictValue
                 Next r
             ElseIf Not IsEmpty(sourceDataArray) Then ' Single row case
                 pqRowCount = 1
                 Dim dictKey As String, dictValue As Variant
                 dictKey = Trim(CStr(tblPQ.ListColumns(PQ_DOCNUM_COL_NAME).DataBodyRange.Value))
                 dictValue = tblPQ.ListColumns(PQ_LOCATION_COL_NAME).DataBodyRange.Value
                 If Len(dictKey) > 0 And Not pqDict.Exists(dictKey) Then pqDict.Add dictKey, dictValue
             Else: pqRowCount = 0
             End If
        End If
    End If
    On Error GoTo LocationErrorHandler ' Restore main handler

    If pqDict.Count > 0 Then Module_Dashboard_UserEdits.LogUserEditsOperation "Loaded " & pqDict.Count & " unique items into lookup dictionary from " & tblPQ.Name
    If pqRowCount = 0 Then Module_Dashboard_UserEdits.LogUserEditsOperation "Warning: Source table '" & tblPQ.Name & "' is empty. Workflow Location will default to 'Quote Only'."

    ' Loop through SORTED Dashboard Rows and Perform Lookup/Write
    Module_Dashboard_UserEdits.LogUserEditsOperation "Matching and writing Workflow Locations row-by-row..."
    For i = 4 To lastRow
        docNum = Trim(CStr(ws.Cells(i, "A").Value))
        locationResult = "Quote Only" ' Default
        If Len(docNum) > 0 Then
            If pqDict.Exists(docNum) Then
                locationResult = pqDict(docNum)
                If IsNull(locationResult) Or IsEmpty(locationResult) Or Len(Trim(CStr(locationResult))) = 0 Then locationResult = "Quote Only"
            End If
        End If
        ws.Cells(i, DB_COL_WORKFLOW_LOCATION).Value = locationResult ' Write directly
    Next i

    Module_Dashboard_UserEdits.LogUserEditsOperation "Finished writing Workflow Locations."
    Debug.Print "PopulateWorkflowLocation (Post-Sort, Row-by-Row Dictionary) Time: " & Timer - t_start
    Set pqDict = Nothing: Exit Sub ' Normal exit

SourceSheetError:
    Module_Dashboard_UserEdits.LogUserEditsOperation "ERROR: Source sheet '" & PQ_LATEST_LOCATION_SHEET & "' not found.": GoTo CleanupExit
SourceTableError:
    Module_Dashboard_UserEdits.LogUserEditsOperation "ERROR: Source table '" & PQ_LATEST_LOCATION_TABLE & "' not found.": GoTo CleanupExit
SourceColumnError:
    Module_Dashboard_UserEdits.LogUserEditsOperation "ERROR: Required columns missing in source table '" & tblPQ.Name & "'.": GoTo CleanupExit
SourceReadError:
    Module_Dashboard_UserEdits.LogUserEditsOperation "ERROR reading data from source table '" & tblPQ.Name & "'.": GoTo CleanupExit
SourceHeaderError:
    Module_Dashboard_UserEdits.LogUserEditsOperation "ERROR: Could not find required headers in source table '" & tblPQ.Name & "'.": GoTo CleanupExit
LocationErrorHandler:
    Module_Dashboard_UserEdits.LogUserEditsOperation "ERROR in PopulateWorkflowLocation: [" & Err.Number & "] " & Err.Description & " (Line: " & Erl & ")"
CleanupExit:
    Set pqDict = Nothing
    Exit Sub
End Sub


Private Sub PopulateMasterQuotesData(ws As Worksheet)
    ' Populates A-I & K with formulas, including static text for K.
    If ws Is Nothing Then Exit Sub
    Dim sourceName As String: sourceName = MASTER_QUOTES_FINAL_SOURCE

    If Not IsMasterQuotesFinalPresent Then
        Module_Dashboard_UserEdits.LogUserEditsOperation "Source '" & sourceName & "' not found. Skipping population of A-I & K."
        Exit Sub
    End If
    On Error GoTo PopulateErrorHandler

    Dim lastMasterRow As Long, sourceTable As ListObject, sourceRange As Range
    Dim isTable As Boolean: isTable = False

    ' Determine source type and row count
    Set sourceTable = Nothing: On Error Resume Next
    Set sourceTable = ws.ListObjects(sourceName) ' Check active sheet first
    If sourceTable Is Nothing Then
        Dim tempWs As Worksheet
        For Each tempWs In ThisWorkbook.Worksheets
            Set sourceTable = tempWs.ListObjects(sourceName): If Not sourceTable Is Nothing Then Exit For
        Next tempWs
    End If
    isTable = Not sourceTable Is Nothing
    On Error GoTo PopulateErrorHandler

    If isTable Then
        If sourceTable.DataBodyRange Is Nothing Then GoTo SourceEmpty
        On Error Resume Next
        lastMasterRow = Application.WorksheetFunction.CountA(sourceTable.ListColumns("Document Number").DataBodyRange)
        If Err.Number <> 0 Or lastMasterRow = 0 Then GoTo SourceEmpty
        On Error GoTo PopulateErrorHandler
    Else ' Assume Named Range
        On Error Resume Next
        Set sourceRange = ThisWorkbook.Names(sourceName).RefersToRange
        If Err.Number <> 0 Or sourceRange Is Nothing Then GoTo SourceInvalid
        On Error GoTo PopulateErrorHandler
        lastMasterRow = Application.WorksheetFunction.CountA(sourceRange.Columns(1))
        If lastMasterRow <= 1 Then GoTo SourceEmpty
        lastMasterRow = lastMasterRow - 1
        sourceName = "'" & sourceRange.Worksheet.Name & "'!" & sourceRange.Address
    End If

    Dim targetRowCount As Long: targetRowCount = lastMasterRow
    If targetRowCount <= 0 Then GoTo SourceEmpty

    Module_Dashboard_UserEdits.LogUserEditsOperation "Populating dashboard columns A-I & K with " & targetRowCount & " rows from " & MASTER_QUOTES_FINAL_SOURCE

    ' Populate with Formulas
    With ws
        Dim formulaBase As String: formulaBase = "=IF(ROWS($A$4:A4)<=" & targetRowCount & ",IFERROR(INDEX(" & sourceName & "[{ColName}],ROWS($A$4:A4)),""""),"""")"
        Dim formula As String
        Dim ifA4NotBlankPrefix As String: ifA4NotBlankPrefix = "=IF(A4<>"""","

        .Range("A4").Resize(targetRowCount, 1).formula = Mid$(Replace(formulaBase, "{ColName}", "Document Number"), 2)
        .Range("B4").Resize(targetRowCount, 1).formula = ifA4NotBlankPrefix & Mid$(Replace(formulaBase, "{ColName}", "Customer Number"), InStr(formulaBase, "IFERROR")) & ","""")"
        .Range("C4").Resize(targetRowCount, 1).formula = ifA4NotBlankPrefix & Mid$(Replace(formulaBase, "{ColName}", "Customer Name"), InStr(formulaBase, "IFERROR")) & ","""")"
        .Range("D4").Resize(targetRowCount, 1).formula = ifA4NotBlankPrefix & Replace(Mid$(Replace(formulaBase, "{ColName}", "Document Amount"), InStr(formulaBase, "IFERROR")), "INDEX", "--INDEX") & ","""")"
        .Range("E4").Resize(targetRowCount, 1).formula = ifA4NotBlankPrefix & Replace(Mid$(Replace(formulaBase, "{ColName}", "Document Date"), InStr(formulaBase, "IFERROR")), "INDEX", "--INDEX") & ","""")"
        .Range("F4").Resize(targetRowCount, 1).formula = ifA4NotBlankPrefix & Replace(Mid$(Replace(formulaBase, "{ColName}", "First Date Pulled"), InStr(formulaBase, "IFERROR")), "INDEX", "--INDEX") & ","""")"
        .Range("G4").Resize(targetRowCount, 1).formula = ifA4NotBlankPrefix & Mid$(Replace(formulaBase, "{ColName}", "Salesperson ID"), InStr(formulaBase, "IFERROR")) & ","""")"
        .Range("H4").Resize(targetRowCount, 1).formula = ifA4NotBlankPrefix & Mid$(Replace(formulaBase, "{ColName}", "User To Enter"), InStr(formulaBase, "IFERROR")) & ","""")"
        .Range("I4").Resize(targetRowCount, 1).formula = ifA4NotBlankPrefix & Mid$(Replace(formulaBase, "{ColName}", "Pull Count"), InStr(formulaBase, "IFERROR")) & ","""")"
        .Range(DB_COL_MISSING_QUOTE & "4").Resize(targetRowCount, 1).formula = ifA4NotBlankPrefix & """Confirm Converted/Voided""" & ","""")" ' Col K Static Text

        ' Format columns
        .Range("D4").Resize(targetRowCount).NumberFormat = "$#,##0.00"
        .Range("E4").Resize(targetRowCount).NumberFormat = "mm/dd/yyyy"
        .Range("F4").Resize(targetRowCount).NumberFormat = "mm/dd/yyyy"
        .Range("I4").Resize(targetRowCount).NumberFormat = "0"
    End With

    Module_Dashboard_UserEdits.LogUserEditsOperation "Finished populating A-I & K."
    Exit Sub
SourceEmpty: Module_Dashboard_UserEdits.LogUserEditsOperation "Source '" & sourceName & "' is empty. Skipping population A-I & K.": Exit Sub
SourceInvalid: Module_Dashboard_UserEdits.LogUserEditsOperation "Source '" & sourceName & "' invalid. Skipping population A-I & K.": Exit Sub
PopulateErrorHandler: Module_Dashboard_UserEdits.LogUserEditsOperation "ERROR in PopulateMasterQuotesData: [" & Err.Number & "] " & Err.Description
End Sub


' --- Layout and Formatting Helpers ---

Private Sub SortDashboardData(ws As Worksheet, lastRow As Long)
    ' Sorts main data range A3:N<lastRow>
    If ws Is Nothing Or lastRow < 5 Then Exit Sub
    Module_Dashboard_UserEdits.LogUserEditsOperation "Sorting dashboard data A3:" & DB_COL_COMMENTS & lastRow
    On Error GoTo SortErrorHandler
    With ws.Sort
        .SortFields.Clear
        .SortFields.Add2 key:=ws.Range("F4:F" & lastRow), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortTextAsNumbers
        .SortFields.Add2 key:=ws.Range("D4:D" & lastRow), SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:=xlSortTextAsNumbers
        .SetRange ws.Range("A3:" & DB_COL_COMMENTS & lastRow)
        .Header = xlYes: .MatchCase = False: .Orientation = xlTopToBottom: .SortMethod = xlPinYin
        .Apply
    End With
    Exit Sub
SortErrorHandler: Module_Dashboard_UserEdits.LogUserEditsOperation "ERROR during sorting: [" & Err.Number & "] " & Err.Description
End Sub

Private Sub FreezeDashboard(ws As Worksheet)
    ' Freezes rows 1-3
    If ws Is Nothing Then Exit Sub
    Dim currentSheetName As String
    On Error Resume Next: currentSheetName = ActiveSheet.Name: On Error GoTo FreezeErrorHandler
    If ws.Name <> currentSheetName Then ws.Activate
    With ActiveWindow
        .FreezePanes = False
        If ws.Cells(ws.Rows.Count, "A").End(xlUp).Row >= 4 Then .FreezePanes = True: ws.Range("A4").Select
    End With
    If ws.Name <> currentSheetName And Len(currentSheetName) > 0 Then On Error Resume Next: ThisWorkbook.Sheets(currentSheetName).Activate: On Error GoTo 0
    Exit Sub
FreezeErrorHandler: Module_Dashboard_UserEdits.LogUserEditsOperation "ERROR applying freeze panes: [" & Err.Number & "] " & Err.Description
End Sub

Private Sub InitializeDashboardLayout(ws As Worksheet)
     ' Sets up rows 1-3 and clears data area A4+
     If ws Is Nothing Then Exit Sub
     Module_Dashboard_UserEdits.LogUserEditsOperation "Initializing dashboard layout for sheet: " & ws.Name
    On Error GoTo InitErrorHandler
    ws.Unprotect ' Ensure unprotected

    ' Setup Rows 1, 2, 3 (Title, Controls, Headers) - Simplified for brevity, see full code above for details
    With ws.Range("A1:" & DB_COL_COMMENTS & "1"): .Merge: .Value = "STRATEGIC QUOTE RECOVERY & CONVERSION TRACKER": .Font.Bold = True: .Interior.Color = RGB(16, 107, 193): .Font.Color = RGB(255, 255, 255): .RowHeight = 32: .HorizontalAlignment = xlCenter: .VerticalAlignment = xlCenter: End With
    With ws.Range("A2:" & DB_COL_COMMENTS & "2"): .Interior.Color = RGB(245, 245, 245): .Borders(xlEdgeTop).LineStyle = xlContinuous: .Borders(xlEdgeBottom).LineStyle = xlContinuous: .RowHeight = 28: End With
    With ws.Range("A2"): .Value = "CONTROL PANEL": .Font.Bold = True: .Interior.Color = RGB(70, 130, 180): .Font.Color = RGB(255, 255, 255): .HorizontalAlignment = xlCenter: .Borders(xlEdgeRight).LineStyle = xlContinuous: End With
    With ws.Range("G2:I2"): .Merge: .Value = "Last Refreshed: Pending...": .HorizontalAlignment = xlCenter: .Font.Size = 9: End With
    With ws.Range(DB_COL_COMMENTS & "2"): .Value = "?": .Font.Bold = True: .Font.Size = 14: .HorizontalAlignment = xlCenter: End With
    With ws.Range("A3:" & DB_COL_COMMENTS & "3"): .Value = Array("Document Number", "Client ID", "Customer Name", "Document Amount", "Document Date", "First Date Pulled", "Salesperson ID", "Entered By", "Pull Count", "Workflow Location", "Missing Quote Alert", "Engagement Phase", "Last Contact Date", "User Comments"): .Font.Bold = True: .Interior.Color = RGB(16, 107, 193): .Font.Color = RGB(255, 255, 255): .HorizontalAlignment = xlCenter: .VerticalAlignment = xlCenter: .WrapText = True: .RowHeight = 30: End With

    ' Clear Data Area & Remove Extra Columns
    ws.Range("A4:" & ws.Cells(ws.Rows.Count, ws.Columns.Count).Address).Clear
    On Error Resume Next: If ws.Columns.Count > 14 Then ws.Range(ws.Columns(15), ws.Columns(ws.Columns.Count)).Delete: On Error GoTo InitErrorHandler

    ' Set Initial Column Widths
    ws.Columns("A").ColumnWidth = 15: ws.Columns("B").ColumnWidth = 12: ws.Columns("C").ColumnWidth = 25: ws.Columns("D").ColumnWidth = 15: ws.Columns("E").ColumnWidth = 12: ws.Columns("F").ColumnWidth = 15: ws.Columns("G").ColumnWidth = 12: ws.Columns("H").ColumnWidth = 15: ws.Columns("I").ColumnWidth = 10: ws.Columns(DB_COL_WORKFLOW_LOCATION).ColumnWidth = 20: ws.Columns(DB_COL_MISSING_QUOTE).ColumnWidth = 25: ws.Columns(DB_COL_PHASE).ColumnWidth = 20: ws.Columns(DB_COL_LASTCONTACT).ColumnWidth = 15: ws.Columns(DB_COL_COMMENTS).ColumnWidth = 40
    ws.Rows("4:" & ws.Rows.Count).RowHeight = 15 ' Default row height

    Exit Sub
InitErrorHandler: Module_Dashboard_UserEdits.LogUserEditsOperation "ERROR in InitializeDashboardLayout: [" & Err.Number & "] " & Err.Description
End Sub

Private Sub CleanupDashboardLayout(ws As Worksheet)
    ' Preserves rows 1-3 if valid, clears A4 down/right.
     If ws Is Nothing Then Exit Sub
    On Error GoTo CleanupErrorHandler
    ws.Unprotect ' Ensure unprotected

    Dim hasTitle As Boolean: hasTitle = (InStr(1, CStr(ws.Range("A1").Value), "STRATEGIC QUOTE RECOVERY", vbTextCompare) > 0)
    Dim hasControls As Boolean: hasControls = (InStr(1, CStr(ws.Range("A2").Value), "CONTROL PANEL", vbTextCompare) > 0)
    Dim hasHeaders As Boolean: hasHeaders = (LCase$(Trim$(CStr(ws.Range("A3").Value))) = "document number")

    If Not hasTitle Or Not hasControls Or Not hasHeaders Then
         Module_Dashboard_UserEdits.LogUserEditsOperation "Core layout rows (1-3) missing/corrupted. Re-initializing layout."
         InitializeDashboardLayout ws ' Rebuilds layout and clears data
    Else
        On Error Resume Next ' Clear data area only
        ws.Range("A4:" & ws.Cells(ws.Rows.Count, ws.Columns.Count).Address).ClearContents
        If Err.Number <> 0 Then Module_Dashboard_UserEdits.LogUserEditsOperation "Warning: Error clearing contents from A4 down: " & Err.Description: Err.Clear
        On Error GoTo CleanupErrorHandler
    End If
    Exit Sub
CleanupErrorHandler: Module_Dashboard_UserEdits.LogUserEditsOperation "ERROR in CleanupDashboardLayout: [" & Err.Number & "] " & Err.Description
End Sub

Private Sub ApplyColorFormatting(ws As Worksheet, lastRow As Long, Optional startDataRow As Long = 4)
    ' Applies conditional formatting to Engagement Phase column (L)
    If ws Is Nothing Then Exit Sub: On Error GoTo FormatErrorHandler
    If lastRow < startDataRow Then Exit Sub ' No data rows
    Dim rngPhase As Range: Set rngPhase = ws.Range(DB_COL_PHASE & startDataRow & ":" & DB_COL_PHASE & lastRow) ' Column L
    ApplyStageFormatting rngPhase ' Call helper sub
    Exit Sub
FormatErrorHandler: Module_Dashboard_UserEdits.LogUserEditsOperation "Error applying Engagement Phase CF: " & Err.Description
End Sub

Private Sub ApplyWorkflowLocationFormatting(ws As Worksheet, lastRow As Long, Optional startDataRow As Long = 4)
    ' Applies conditional formatting to Workflow Location column (J)
    If ws Is Nothing Then Exit Sub: On Error GoTo FormatErrorHandler
    If lastRow < startDataRow Then Exit Sub ' No data rows
    Dim rngLocation As Range: Set rngLocation = ws.Range(DB_COL_WORKFLOW_LOCATION & startDataRow & ":" & DB_COL_WORKFLOW_LOCATION & lastRow)
    rngLocation.FormatConditions.Delete ' Clear existing
    Dim fc As FormatCondition
    With rngLocation ' Add rules using constants for values
        Set fc = .FormatConditions.Add(xlCellValue, xlEqual, "=""Quote Only"""): fc.Interior.Color = RGB(230, 240, 248)
        Set fc = .FormatConditions.Add(xlCellValue, xlEqual, "=""1. New Orders"""): fc.Interior.Color = RGB(208, 230, 245)
        Set fc = .FormatConditions.Add(xlCellValue, xlEqual, "=""2. Open Orders"""): fc.Interior.Color = RGB(255, 247, 209)
        Set fc = .FormatConditions.Add(xlCellValue, xlEqual, "=""3. As Available Orders"""): fc.Interior.Color = RGB(227, 215, 232)
        Set fc = .FormatConditions.Add(xlCellValue, xlEqual, "=""4. Hold Orders"""): fc.Interior.Color = RGB(255, 150, 54)
        Set fc = .FormatConditions.Add(xlCellValue, xlEqual, "=""5. Closed Files"""): fc.Interior.Color = RGB(230, 230, 230)
        Set fc = .FormatConditions.Add(xlCellValue, xlEqual, "=""6. Declined Orders"""): fc.Interior.Color = RGB(209, 47, 47)
    End With
    Exit Sub
FormatErrorHandler: Module_Dashboard_UserEdits.LogUserEditsOperation "Error applying Workflow Location CF: " & Err.Description
End Sub

Private Sub ApplyStageFormatting(targetRng As Range)
    ' Helper sub - Contains the CF rules for Engagement Phase (Column L)
    If targetRng Is Nothing Or targetRng.Cells.CountLarge = 0 Then Exit Sub
    On Error GoTo HelperErrorHandler
    targetRng.FormatConditions.Delete ' Clear existing
    Dim fc As FormatCondition
    Const compType As XlFormatConditionType = xlCellValue
    Const compOperator As XlFormatConditionOperator = xlEqual
    Const formulaPrefix As String = "="""
    Const formulaSuffix As String = """"
    Const STOP_IF_TRUE As Boolean = True

    With targetRng ' Add rules, highest priority first
        Set fc = .FormatConditions.Add(compType, compOperator, formulaPrefix & "Converted" & formulaSuffix): With fc.Interior: .Color = RGB(120, 235, 120): End With: With fc.Font: .Bold = True: End With: fc.StopIfTrue = STOP_IF_TRUE
        Set fc = .FormatConditions.Add(compType, compOperator, formulaPrefix & "Declined" & formulaSuffix): With fc.Interior: .Color = RGB(209, 47, 47): End With: fc.StopIfTrue = STOP_IF_TRUE
        Set fc = .FormatConditions.Add(compType, compOperator, formulaPrefix & "Closed (Extra Order)" & formulaSuffix): With fc.Interior: .Color = RGB(184, 39, 39): End With: fc.StopIfTrue = STOP_IF_TRUE
        Set fc = .FormatConditions.Add(compType, compOperator, formulaPrefix & "Closed" & formulaSuffix): With fc.Interior: .Color = RGB(166, 28, 28): End With: fc.StopIfTrue = STOP_IF_TRUE
        ' Non-outcome statuses (StopIfTrue=False)
        Set fc = .FormatConditions.Add(compType, compOperator, formulaPrefix & "First F/U" & formulaSuffix): With fc.Interior: .Color = RGB(208, 230, 245): End With
        Set fc = .FormatConditions.Add(compType, compOperator, formulaPrefix & "Second F/U" & formulaSuffix): With fc.Interior: .Color = RGB(146, 198, 237): End With
        Set fc = .FormatConditions.Add(compType, compOperator, formulaPrefix & "Third F/U" & formulaSuffix): With fc.Interior: .Color = RGB(245, 225, 113): End With
        Set fc = .FormatConditions.Add(compType, compOperator, formulaPrefix & "Long-Term F/U" & formulaSuffix): With fc.Interior: .Color = RGB(255, 150, 54): End With
        Set fc = .FormatConditions.Add(compType, compOperator, formulaPrefix & "Requoting" & formulaSuffix): With fc.Interior: .Color = RGB(227, 215, 232): End With
        Set fc = .FormatConditions.Add(compType, compOperator, formulaPrefix & "Pending" & formulaSuffix): With fc.Interior: .Color = RGB(255, 247, 209): End With
        Set fc = .FormatConditions.Add(compType, compOperator, formulaPrefix & "No Response" & formulaSuffix): With fc.Interior: .Color = RGB(245, 238, 224): End With
        Set fc = .FormatConditions.Add(compType, compOperator, formulaPrefix & "Texas (No F/U)" & formulaSuffix): With fc.Interior: .Color = RGB(230, 217, 204): End With
        Set fc = .FormatConditions.Add(compType, compOperator, formulaPrefix & "AF" & formulaSuffix): With fc.Interior: .Color = RGB(162, 217, 210): End With
        Set fc = .FormatConditions.Add(compType, compOperator, formulaPrefix & "RZ" & formulaSuffix): With fc.Interior: .Color = RGB(138, 155, 212): End With
        Set fc = .FormatConditions.Add(compType, compOperator, formulaPrefix & "KMH" & formulaSuffix): With fc.Interior: .Color = RGB(247, 196, 175): End With
        Set fc = .FormatConditions.Add(compType, compOperator, formulaPrefix & "RI" & formulaSuffix): With fc.Interior: .Color = RGB(191, 225, 243): End With
        Set fc = .FormatConditions.Add(compType, compOperator, formulaPrefix & "WW/OM" & formulaSuffix): With fc.Interior: .Color = RGB(155, 124, 185): End With
    End With
    Exit Sub
HelperErrorHandler: Module_Dashboard_UserEdits.LogUserEditsOperation "Error applying stage formatting rules: " & Err.Description
End Sub

Private Sub ProtectUserColumns(ws As Worksheet)
    ' Locks columns except L:N, protects sheet.
    If ws Is Nothing Then Exit Sub
    On Error GoTo ProtectErrorHandler
    ws.Unprotect ' Unprotect first
    ws.Cells.Locked = True ' Lock all
    Dim lastRowProt As Long: lastRowProt = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    If lastRowProt >= 4 Then On Error Resume Next: ws.Range(DB_COL_PHASE & "4:" & DB_COL_COMMENTS & lastRowProt).Locked = False: On Error GoTo ProtectErrorHandler ' Unlock L:N
    ws.Protect UserInterfaceOnly:=True, DrawingObjects:=True, Contents:=True, Scenarios:=True ' Re-protect
    Exit Sub
ProtectErrorHandler: Module_Dashboard_UserEdits.LogUserEditsOperation "ERROR applying protection: " & Err.Description
End Sub

Private Sub ModernButton(ws As Worksheet, cellRef As String, buttonText As String, macroName As String)
    ' Creates styled buttons in row 2
    If ws Is Nothing Then Exit Sub
    Dim btn As Shape, targetCell As Range, buttonTop As Double, buttonLeft As Double, buttonWidth As Double, buttonHeight As Double
    On Error GoTo ButtonErrorHandler
    Set targetCell = ws.Range(cellRef): If targetCell Is Nothing Then Exit Sub
    buttonLeft = targetCell.Left + 2: buttonTop = targetCell.Top + 2
    buttonWidth = Application.Max(targetCell.MergeArea.Width - 4, 60)
    buttonHeight = targetCell.Height - 4
    On Error Resume Next: Set btn = ws.Shapes.AddShape(msoShapeRoundedRectangle, buttonLeft, buttonTop, buttonWidth, buttonHeight): If Err.Number <> 0 Then Exit Sub
    On Error GoTo ButtonErrorHandler
    With btn ' Style the button
        With .TextFrame2.TextRange: .Text = buttonText: With .Font: .Fill.ForeColor.RGB = RGB(255, 255, 255): .Size = 10: .Name = "Segoe UI": .Bold = msoTrue: End With: End With
        .TextFrame2.HorizontalAnchor = msoAnchorCenter: .TextFrame2.VerticalAnchor = msoAnchorMiddle: .TextFrame2.WordWrap = msoFalse: .TextFrame2.AutoSize = msoAutoSizeNone
        .Fill.Solid: .Fill.ForeColor.RGB = RGB(42, 120, 180): .Line.Visible = msoTrue: .Line.ForeColor.RGB = RGB(25, 95, 150): .Line.Weight = 0.75
        On Error Resume Next: .Shadow.Type = msoShadow21: .Shadow.Transparency = 0.7: .Shadow.Visible = msoTrue: On Error GoTo ButtonErrorHandler
        .OnAction = macroName: .LockAspectRatio = msoFalse
    End With
    Exit Sub
ButtonErrorHandler: Module_Dashboard_UserEdits.LogUserEditsOperation "Error in ModernButton: " & Err.Description
End Sub

Private Function GetOrCreateDashboardSheet(sheetName As String) As Worksheet
    ' Returns or creates the SQRCT Dashboard sheet
    Dim ws As Worksheet
    On Error Resume Next: Set ws = ThisWorkbook.Sheets(sheetName): Dim sheetExists As Boolean: sheetExists = (Err.Number = 0 And Not ws Is Nothing): Err.Clear: On Error GoTo 0
    If Not sheetExists Then
        On Error Resume Next: Set ws = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count)): If Err.Number <> 0 Then Exit Function ' Failed to add
        ws.Name = sheetName: If Err.Number <> 0 Then Err.Clear ' Use default name if rename fails
        On Error GoTo 0: SetupDashboard ws ' Call setup for new sheet
    End If
    Set GetOrCreateDashboardSheet = ws
End Function

Private Function IsMasterQuotesFinalPresent() As Boolean
    ' Checks if the primary data source exists (PQ, Table, or Named Range)
    Dim lo As ListObject, nm As Name, queryObj As Object, tempRange As Range
    On Error Resume Next ' Ignore errors for checking different types
    Set queryObj = ThisWorkbook.Queries(MASTER_QUOTES_FINAL_SOURCE): If Err.Number = 0 And Not queryObj Is Nothing Then IsMasterQuotesFinalPresent = True: GoTo FoundIt
    Set lo = Nothing: Dim ws As Worksheet: For Each ws In ThisWorkbook.Worksheets: Set lo = ws.ListObjects(MASTER_QUOTES_FINAL_SOURCE): If Err.Number = 0 And Not lo Is Nothing Then IsMasterQuotesFinalPresent = True: GoTo FoundIt: Err.Clear: Next ws
    Set nm = ThisWorkbook.Names(MASTER_QUOTES_FINAL_SOURCE): If Err.Number = 0 And Not nm Is Nothing Then Set tempRange = nm.RefersToRange: If Not tempRange Is Nothing Then IsMasterQuotesFinalPresent = True: GoTo FoundIt
FoundIt: On Error GoTo 0
End Function

Private Sub CreateOrUpdateTextOnlySheet(wsSource As Worksheet)
    ' Helper to manage the Text-Only sheet creation/update
    If wsSource Is Nothing Then Exit Sub
    Dim wsValues As Worksheet, currentSheet As Worksheet, lastRowSource As Long, lastRowValues As Long, srcRange As Range
    Const TEXT_ONLY_SHEET_NAME As String = "SQRCT Dashboard (Text-Only)"
    On Error GoTo TextOnlyErrorHandler
    Set currentSheet = ActiveSheet

    ' Get or Create Sheet
    On Error Resume Next: Set wsValues = ThisWorkbook.Sheets(TEXT_ONLY_SHEET_NAME): On Error GoTo TextOnlyErrorHandler
    If wsValues Is Nothing Then On Error Resume Next: Set wsValues = ThisWorkbook.Sheets.Add(After:=wsSource): If Err.Number = 0 Then wsValues.Name = TEXT_ONLY_SHEET_NAME: If Err.Number <> 0 Then GoTo TextOnlyErrorHandler: On Error GoTo TextOnlyErrorHandler: Module_Dashboard_UserEdits.LogUserEditsOperation "Created sheet: " & wsValues.Name
    Else: wsValues.Visible = xlSheetVisible: wsValues.Cells.Clear: If Err.Number <> 0 Then Module_Dashboard_UserEdits.LogUserEditsOperation "Warning: Error clearing Text-Only sheet.": Err.Clear: Module_Dashboard_UserEdits.LogUserEditsOperation "Cleared existing sheet: " & TEXT_ONLY_SHEET_NAME
    End If

    ' Copy Data
    lastRowSource = wsSource.Cells(wsSource.Rows.Count, "A").End(xlUp).Row
    If lastRowSource >= 3 Then Set srcRange = wsSource.Range("A3:" & DB_COL_COMMENTS & lastRowSource): srcRange.Copy: wsValues.Range("A1").PasteSpecial Paste:=xlPasteValuesAndNumberFormats: Application.CutCopyMode = False: Module_Dashboard_UserEdits.LogUserEditsOperation "Pasted data to " & TEXT_ONLY_SHEET_NAME

    ' Apply CF
    lastRowValues = wsValues.Cells(wsValues.Rows.Count, "A").End(xlUp).Row
    If lastRowValues >= 2 Then ApplyColorFormatting wsValues, lastRowValues, 2: ApplyWorkflowLocationFormatting wsValues, lastRowValues, 2: Module_Dashboard_UserEdits.LogUserEditsOperation "Applied CF to " & TEXT_ONLY_SHEET_NAME

    ' Final Formatting
    wsValues.Columns("A:" & DB_COL_COMMENTS).AutoFit
    If wsValues.Columns("N").ColumnWidth < 25 Then wsValues.Columns("N").ColumnWidth = 25 ' Min width Comments
    On Error Resume Next: wsValues.Unprotect: If ActiveSheet.Name = wsValues.Name Then ActiveWindow.FreezePanes = False: On Error GoTo TextOnlyErrorHandler
    If ActiveSheet.Name <> currentSheet.Name Then currentSheet.Activate

    Exit Sub
TextOnlyErrorHandler: Module_Dashboard_UserEdits.LogUserEditsOperation "ERROR creating/updating Text-Only sheet: " & Err.Description: Application.CutCopyMode = False: On Error Resume Next: If Not currentSheet Is Nothing Then currentSheet.Activate: On Error GoTo 0
End Sub



