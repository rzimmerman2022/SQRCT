'=====================================================================
'  Module :  modArchival  – Archive / Active / All dashboard views
'  Workbook: SQRCT
'---------------------------------------------------------------------
'  Purpose :
'  •  Provide three read-only snapshots of the main dashboard:
'       –  "All"      = normal dashboard (no filter)
'       –  "Active"   = only rows still in Client-Relations scope
'       –  "Archive"  = everything else (legacy, completed, OM scope)
'  •  Drop navigation buttons on every view so the user can hop between
'     sheets without hunting the tabs.
'  •  Keep every public entry-point self-contained so Module_Dashboard can
'     simply `Call modArchival.RefreshAllViews`.
'=====================================================================
Option Explicit

'--------------------------------------------------
' 1--  PHASE CATEGORISATION  (Based on user confirmation)
'--------------------------------------------------
' Uses Pipe Delimiter for efficient InStr check (|Value|)
' Phases NOT listed here will be considered Archived.
Private Const ACTIVE_PHASES As String = "|First F/U|AF|RZ|KMH|RI|" ' Added leading/trailing pipes

'--------------------------------------------------
' 2--  SHEET NAMES & UI TEXT
'--------------------------------------------------
' Using SH_DASH from Module_Dashboard via Public Const
Private Const SH_ACTIVE As String = "SQRCT Active"     ' filtered view name
Private Const SH_ARCHIVE As String = "SQRCT Archive"    ' filtered view name
Private Const TITLE_ACTIVE As String = _
        "STRATEGIC QUOTE RECOVERY & CONVERSION TRACKER – ACTIVE VIEW" ' Sheet title
Private Const TITLE_ARCHIVE As String = _
        "STRATEGIC QUOTE RECOVERY & CONVERSION TRACKER – ARCHIVE VIEW" ' Sheet title

'--------------------------------------------------
' 3--  PUBLIC BUTTON WRAPPERS  (Called by buttons)
'--------------------------------------------------
Public Sub btnViewActive()
    ' Refreshes the view THEN activates it
    Application.StatusBar = "Refreshing Active View..."
    RefreshAndActivate SH_ACTIVE
    Application.StatusBar = False
End Sub

Public Sub btnViewArchive()
    ' Refreshes the view THEN activates it
    Application.StatusBar = "Refreshing Archive View..."
    RefreshAndActivate SH_ARCHIVE
    Application.StatusBar = False
End Sub

Public Sub btnViewAll()
    ' Just activates the main dashboard
    Application.StatusBar = "Activating Main Dashboard..."
    On Error Resume Next
    ThisWorkbook.Worksheets(Module_Dashboard.DASHBOARD_SHEET_NAME).Activate
    If Err.Number <> 0 Then MsgBox "Could not activate sheet: " & Module_Dashboard.DASHBOARD_SHEET_NAME, vbExclamation
    Application.StatusBar = False
End Sub

'--------------------------------------------------
' 4--  TOP-LEVEL REFRESH ENTRY (Called by Module_Dashboard.RefreshDashboard)
'--------------------------------------------------
Public Sub RefreshAllViews()
    Dim wsDash As Worksheet
    Dim originalCalcState As XlCalculation
    Dim originalEventsState As Boolean
    Dim originalScreenState As Boolean

    Log "[Archival] RefreshAllViews – start"

    ' --- Store original settings ---
    originalCalcState = Application.Calculation
    originalEventsState = Application.EnableEvents
    originalScreenState = Application.ScreenUpdating

    ' --- Apply performance settings ---
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    Application.StatusBar = "Refreshing Active/Archive views..."

    ' --- Get Main Dashboard Reference ---
    On Error Resume Next
    Set wsDash = ThisWorkbook.Worksheets(Module_Dashboard.DASHBOARD_SHEET_NAME)
    On Error GoTo ArchivalErrorHandler ' Use a specific handler for this sub

    If wsDash Is Nothing Then
        Log "[Archival] ERROR: Main dashboard sheet '" & Module_Dashboard.DASHBOARD_SHEET_NAME & "' not found. Cannot refresh views."
        GoTo ArchivalCleanup ' Go to cleanup
    End If

    ' --- Refresh Views ---
    RefreshActiveView wsDash ' Pass source sheet
    RefreshArchiveView wsDash ' Pass source sheet
    AddNavigationButtons wsDash ' Add buttons to main dashboard

ArchivalCleanup: ' Cleanup label for both normal exit and error
    On Error Resume Next ' Ignore errors during cleanup
    Log "[Archival] RefreshAllViews – cleaning up"
    ' --- Restore original settings ---
    Application.Calculation = originalCalcState
    Application.EnableEvents = originalEventsState
    Application.ScreenUpdating = originalScreenState
    Application.StatusBar = False ' Clear status bar
    Set wsDash = Nothing
    Log "[Archival] RefreshAllViews – done"
    Exit Sub

ArchivalErrorHandler: ' Error handler for this sub
    Log "[Archival] ERROR in RefreshAllViews [" & Err.Number & "]: " & Err.Description
    MsgBox "An error occurred while refreshing the Active/Archive views:" & vbCrLf & Err.Description, vbCritical, "Archival Refresh Error"
    Resume ArchivalCleanup ' Go to cleanup after logging error

End Sub

'*********************** PRIVATE SECTION  ***************************

'--------------------------------------------------
'  Refresh & activate a single view (helper)
'--------------------------------------------------
Private Sub RefreshAndActivate(viewName As String)
    Dim wsDash As Worksheet
    Dim originalCalcState As XlCalculation
    Dim originalEventsState As Boolean
    Dim originalScreenState As Boolean

    Log "[Archival] RefreshAndActivate: Starting for '" & viewName & "'"

    ' --- Store original settings ---
    originalCalcState = Application.Calculation
    originalEventsState = Application.EnableEvents
    originalScreenState = Application.ScreenUpdating

    ' --- Apply performance settings ---
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    Application.StatusBar = "Refreshing " & viewName & "..." ' Update status

    ' --- Get Main Dashboard Reference ---
    On Error Resume Next
    Set wsDash = ThisWorkbook.Worksheets(Module_Dashboard.DASHBOARD_SHEET_NAME)
    On Error GoTo RefreshActivateErrorHandler ' Specific handler

    If wsDash Is Nothing Then
        Log "[Archival] ERROR in RefreshAndActivate: Main dashboard sheet not found."
        GoTo RefreshActivateCleanup ' Go to cleanup
    End If

    ' --- Refresh the specific view ---
    Select Case viewName
        Case SH_ACTIVE: RefreshActiveView wsDash
        Case SH_ARCHIVE: RefreshArchiveView wsDash
        Case Else
            Log "[Archival] ERROR in RefreshAndActivate: Invalid view name '" & viewName & "'."
            GoTo RefreshActivateCleanup
    End Select

    ' --- Activate the sheet ---
    On Error Resume Next ' Ignore error activating (sheet might be hidden temporarily)
    ThisWorkbook.Worksheets(viewName).Activate
    If Err.Number <> 0 Then
        Log "[Archival] ERROR in RefreshAndActivate: Could not activate sheet '" & viewName & "'. Error: " & Err.Description
        Err.Clear ' Clear the error
    End If
    On Error GoTo RefreshActivateErrorHandler ' Restore handler

RefreshActivateCleanup:
    On Error Resume Next ' Ignore errors during cleanup
    Log "[Archival] RefreshAndActivate: Cleaning up for '" & viewName & "'"
    ' --- Restore original settings ---
    Application.Calculation = originalCalcState
    Application.EnableEvents = originalEventsState
    Application.ScreenUpdating = originalScreenState
    Application.StatusBar = False ' Clear status bar
    Set wsDash = Nothing
    Log "[Archival] RefreshAndActivate: Done for '" & viewName & "'"
    Exit Sub

RefreshActivateErrorHandler:
     Log "[Archival] ERROR in RefreshAndActivate for '" & viewName & "' [" & Err.Number & "]: " & Err.Description
     MsgBox "An error occurred while refreshing/activating the " & viewName & " view:" & vbCrLf & Err.Description, vbExclamation, "View Activation Error"
     Resume RefreshActivateCleanup ' Go to cleanup

End Sub

'--------------------------------------------------
'  ACTIVE  VIEW  (filter-in rows *where phase is active*)
'--------------------------------------------------
Private Sub RefreshActiveView(wsSrc As Worksheet) ' Takes source ws as parameter
    Dim wsTgt As Worksheet
    Log "[Archival] RefreshActiveView: Starting..."
    Set wsTgt = GetOrCreateSheet(SH_ACTIVE, TITLE_ACTIVE, RGB(0, 110, 0)) ' Dark Green Banner

    If wsTgt Is Nothing Then
        Log "[Archival] RefreshActiveView: Failed to get/create target sheet. Aborting."
        Exit Sub
    End If

    ' Filter condition: Keep if Phase is Active (IsPhaseArchived = False)
    CopyFilteredRows wsSrc, wsTgt, keepArchived:=False
    ApplyViewFormatting wsTgt, "Active" ' Apply formatting and buttons
    Log "[Archival] RefreshActiveView: Completed."
End Sub

'--------------------------------------------------
'  ARCHIVE VIEW  (filter-in rows *where phase is archived*)
'--------------------------------------------------
Private Sub RefreshArchiveView(wsSrc As Worksheet) ' Takes source ws as parameter
    Dim wsTgt As Worksheet
    Log "[Archival] RefreshArchiveView: Starting..."
    Set wsTgt = GetOrCreateSheet(SH_ARCHIVE, TITLE_ARCHIVE, RGB(150, 40, 40)) ' Dark Red Banner

    If wsTgt Is Nothing Then
        Log "[Archival] RefreshArchiveView: Failed to get/create target sheet. Aborting."
        Exit Sub
    End If

    ' Filter condition: Keep if Phase is Archived (IsPhaseArchived = True)
    CopyFilteredRows wsSrc, wsTgt, keepArchived:=True
    ApplyViewFormatting wsTgt, "Archive" ' Apply formatting and buttons
    Log "[Archival] RefreshArchiveView: Completed."
End Sub

'--------------------------------------------------
'  Copy rows from main dashboard based on phase test (Uses Array)
'--------------------------------------------------
Private Sub CopyFilteredRows(wsSrc As Worksheet, wsTgt As Worksheet, _
                             keepArchived As Boolean)
    Dim lastSrcRow As Long, destRow As Long, r As Long
    Dim phaseValue As String
    Dim arrData As Variant      ' Array to hold source data A:N
    Dim i As Long, j As Long    ' Loop counters
    Dim outputData() As Variant ' Array to hold rows to be copied
    Dim outputRowCount As Long
    Dim srcRange As Range
    Dim phaseColIndex As Long
    Dim numCols As Long

    ' --- Define Source Data Range (A to N) ---
    ' Use Public Constants from Module_Dashboard
    Dim lastColLetter As String: lastColLetter = Module_Dashboard.DB_COL_COMMENTS ' Should be "N"
    Dim firstDataRow As Long: firstDataRow = 4 ' Data starts at row 4 on main dash

    lastSrcRow = wsSrc.Cells(wsSrc.rows.Count, "A").End(xlUp).Row
    ' numCols = wsSrc.Columns(lastColLetter).Column ' Remove

    Log "[Archival] CopyFilteredRows: Reading source '" & wsSrc.Name & "' rows " & firstDataRow & " to " & lastSrcRow & "."

    ' --- Handle No Data Case ---
    If lastSrcRow < firstDataRow Then
        Log "[Archival] CopyFilteredRows: No data rows found on source sheet '" & wsSrc.Name & "'."
        wsTgt.Cells.ClearContents ' Ensure target is clear
        wsSrc.Range("A1:" & lastColLetter & "3").Copy wsTgt.Range("A1") ' Copy A1:N3 (Title/Control/Headers)
        ' Add count below timestamp area
        On Error Resume Next
        With wsTgt.Range("G2:I2") ' Assuming timestamp area GHI
             If Not .MergeCells Then .Merge ' Ensure merged
             .Value = vbNullString ' Clear timestamp
             .HorizontalAlignment = xlCenter
             .VerticalAlignment = xlCenter
             .Cells(1, .Cells.Count).Offset(0, 1).Value = IIf(keepArchived, "Archive", "Active") & " Count: 0" ' Put count in J2
             .Cells(1, .Cells.Count).Offset(0, 1).ClearFormats
             .Cells(1, .Cells.Count).Offset(0, 1).HorizontalAlignment = xlLeft
             .Cells(1, .Cells.Count).Offset(0, 1).VerticalAlignment = xlCenter
             .Cells(1, .Cells.Count).Offset(0, 1).Font.Size = 9
        End With
        On Error GoTo 0
        Exit Sub
    End If

    ' --- Read source data into array for faster processing ---
    On Error Resume Next
    Set srcRange = wsSrc.Range("A" & firstDataRow & ":" & lastColLetter & lastSrcRow) ' A4:N<lastRow>
    ' ...(Code that handles reading range into arrData, might include If/Else for single cell)...
    arrData = srcRange.Value2 ' This line (or similar) reads the data

    ' Check if reading failed or didn't produce an array
    If Err.Number <> 0 Or Not IsArray(arrData) Then
        Log "[Archival] CopyFilteredRows: Error reading source data into array. " & Err.Description
        Exit Sub
    End If
    On Error GoTo 0 ' Restore default handler

    ' --- Get TRUE column count from the array Excel returned --- <<< PASTE IT HERE
    numCols = UBound(arrData, 2)

    ' --- (Optional) Add check/log if count differs ---
    If numCols <> wsSrc.Columns(lastColLetter).Column Then
        Log "[Archival] CopyFilteredRows: Warning - Source data array only contains " & numCols & _
            " columns (expected " & wsSrc.Columns(lastColLetter).Column & "). Trailing empty columns likely omitted."
    End If

    ' --- Handle Single Row Case (Value2 read gives 1D array) ---
    ' ...(Code continues)...

    If Err.Number <> 0 Or Not IsArray(arrData) Then
        Log "[Archival] CopyFilteredRows: Error reading source data into array. " & Err.Description
        Exit Sub
    End If
    On Error GoTo 0 ' Restore default handler

    ' --- Handle Single Row Case (Value2 read gives 1D array) ---
    If lastSrcRow = firstDataRow Then
        ' Check if it's not already a 2D array by checking bounds
         Dim is2D As Boolean
         On Error Resume Next
         Dim dummy As Long: dummy = UBound(arrData, 2)
         is2D = (Err.Number = 0)
         On Error GoTo 0

         If Not is2D Then ' It's 1D
             Log "[Archival] CopyFilteredRows: Converting single row 1D data to 2D array."
             Dim tempArr() As Variant
             ReDim tempArr(1 To 1, 1 To UBound(arrData, 1)) ' Size based on 1D upper bound
              If UBound(tempArr, 2) <> numCols Then
                 Log "[Archival] CopyFilteredRows: Warning - Single row data item count (" & UBound(tempArr, 2) & ") doesn't match expected columns (" & numCols & ")."
                 ReDim tempArr(1 To 1, 1 To numCols) ' Resize to expected columns
             End If

             For j = LBound(arrData, 1) To UBound(arrData, 1)
                  If j <= numCols Then tempArr(1, j) = arrData(j) ' Avoid error if 1D > numCols
             Next j
             Erase arrData
             arrData = tempArr ' Replace original 1D array with 2D
         End If
    End If


    ' --- Prepare output array ---
    ReDim outputData(1 To UBound(arrData, 1), 1 To numCols) ' Max possible size
    outputRowCount = 0

    ' --- Get the 1-based index for the Phase column (L) within the array ---
    On Error Resume Next ' In case constant refers to invalid column
    phaseColIndex = wsSrc.Columns(Module_Dashboard.DB_COL_PHASE).Column
    If Err.Number <> 0 Or phaseColIndex = 0 Or phaseColIndex > numCols Then
        Log "[Archival] CopyFilteredRows: ERROR - Invalid Phase Column '" & Module_Dashboard.DB_COL_PHASE & "' index (" & phaseColIndex & ")."
        Erase arrData ' Clean up input array
        Exit Sub
    End If
    On Error GoTo 0

    ' --- Filter rows in memory ---
    Log "[Archival] CopyFilteredRows: Filtering " & UBound(arrData, 1) & " rows in memory..."
    For r = LBound(arrData, 1) To UBound(arrData, 1) ' Loop through rows in the source array
        ' Get phase from array using the calculated column index
        phaseValue = Trim$(CStr(arrData(r, phaseColIndex)))

        ' Check if the row's status matches the filter criteria
        ' KeepArchived = TRUE means we keep rows where IsPhaseArchived is TRUE
        ' KeepArchived = FALSE means we keep rows where IsPhaseArchived is FALSE (i.e., Active)
        If IsPhaseArchived(phaseValue) = keepArchived Then
            ' Row matches criteria, add to output array
            outputRowCount = outputRowCount + 1
            For i = 1 To numCols ' Copy all columns (A-N)
                outputData(outputRowCount, i) = arrData(r, i)
            Next i
        End If
    Next r
    Log "[Archival] CopyFilteredRows: Found " & outputRowCount & " matching rows."

    ' --- Write results to Target Sheet ---
    Application.EnableEvents = False ' Keep events off during write
    wsTgt.Cells.ClearContents                       ' Clear target completely first
    wsSrc.Range("A1:" & lastColLetter & "3").Copy wsTgt.Range("A1") ' Copy A1:N3 (Title/Control/Headers)

    If outputRowCount > 0 Then
        ' Resize output array to actual size before writing
        ReDim Preserve outputData(1 To outputRowCount, 1 To numCols)
        ' Write filtered data starting at A4
        On Error Resume Next
        wsTgt.Range("A4").Resize(outputRowCount, numCols).Value = outputData
        If Err.Number <> 0 Then
             Log "[Archival] CopyFilteredRows: ERROR writing data array to '" & wsTgt.Name & "'. " & Err.Description
        Else
             Log "[Archival] CopyFilteredRows: Wrote " & outputRowCount & " rows to '" & wsTgt.Name & "'."
        End If
        On Error GoTo 0
    Else
        Log "[Archival] CopyFilteredRows: No rows matched filter criteria for '" & wsTgt.Name & "'. Sheet is empty except headers."
        ' Ensure data area is clear if no rows written
        wsTgt.Range("A4:" & lastColLetter & wsTgt.rows.Count).ClearContents
    End If

    ' Add count below timestamp area (handle potential merge)
    On Error Resume Next
    With wsTgt.Range("G2:I2") ' Assuming timestamp area GHI
         If Not .MergeCells Then .Merge ' Ensure merged
         .Value = vbNullString ' Clear any old timestamp value
         .HorizontalAlignment = xlCenter
         .VerticalAlignment = xlCenter
         .Cells(1, .Cells.Count).Offset(0, 1).Value = IIf(keepArchived, "Archive", "Active") & " Count: " & outputRowCount ' Put count in J2
         .Cells(1, .Cells.Count).Offset(0, 1).ClearFormats
         .Cells(1, .Cells.Count).Offset(0, 1).HorizontalAlignment = xlLeft
         .Cells(1, .Cells.Count).Offset(0, 1).VerticalAlignment = xlCenter
         .Cells(1, .Cells.Count).Offset(0, 1).Font.Size = 9
    End With
    On Error GoTo 0

    Application.EnableEvents = True ' Re-enable events

    ' Clean up memory
    Erase arrData
    Erase outputData

End Sub

'--------------------------------------------------
'  Phase test helper - Checks if phase is considered Active
'--------------------------------------------------
Private Function IsPhaseActive(ph As String) As Boolean
    ' Returns TRUE if the phase is found within the ACTIVE_PHASES constant
    If Len(ph) = 0 Then Exit Function ' Blank is not active
    ' Uses pipe delimiters for exact match and case-insensitive compare
    IsPhaseActive = (InStr(1, ACTIVE_PHASES, "|" & ph & "|", vbTextCompare) > 0)
End Function

'--------------------------------------------------
'  Phase test helper - Checks if phase is considered Archived
'--------------------------------------------------
Private Function IsPhaseArchived(ph As String) As Boolean
    ' A phase is considered archived if it's NOT blank and NOT in the Active list
    ' Also treats blank phases as archived.
    If Len(ph) = 0 Then
        IsPhaseArchived = True ' Treat blank phase as archived/out-of-scope
    Else
        IsPhaseArchived = Not IsPhaseActive(ph) ' Archived = Not Active
    End If
End Function

'--------------------------------------------------
'  Create sheet if missing + drop coloured title bar
'--------------------------------------------------
Private Function GetOrCreateSheet(sName As String, sTitle As String, _
                                  bannerColor As Long) As Worksheet
    Dim ws As Worksheet
    Dim wbStructureWasLocked As Boolean

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(sName)
    On Error GoTo GetOrCreateSheet_Error ' Use specific handler

    If ws Is Nothing Then
        Log "[Archival] GetOrCreateSheet: Sheet '" & sName & "' not found. Attempting creation..."
        ' --- Need to temporarily unprotect workbook structure to add sheet ---
        wbStructureWasLocked = ThisWorkbook.ProtectStructure
        If wbStructureWasLocked Then
            If Not Module_Dashboard.ToggleWorkbookStructure(False) Then ' Use Toggle function from Module_Dashboard
                Log "[Archival] GetOrCreateSheet: FATAL - Failed to unprotect workbook structure to add sheet '" & sName & "'."
                Exit Function ' Cannot proceed
            End If
        End If
        ' --- Add the sheet ---
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        If ws Is Nothing Then ' Check if Add failed
             Log "[Archival] GetOrCreateSheet: FATAL - Failed to add new sheet."
             ' Attempt to re-lock structure if we unlocked it
             If wbStructureWasLocked Then Module_Dashboard.ToggleWorkbookStructure (True)
             Exit Function
        End If

        ' --- Name the sheet ---
        On Error Resume Next ' Handle error during Name
        ws.Name = sName
        If Err.Number <> 0 Then
             Log "[Archival] GetOrCreateSheet: WARNING - Failed to name sheet '" & sName & "'. Using default '" & ws.Name & "'. Error: " & Err.Description
             Err.Clear
        End If
        On Error GoTo GetOrCreateSheet_Error ' Restore handler

         ' --- Re-lock structure if needed ---
         If wbStructureWasLocked Then
             If Not Module_Dashboard.ToggleWorkbookStructure(True) Then
                 Log "[Archival] GetOrCreateSheet: WARNING - Failed to re-protect workbook structure after adding sheet '" & ws.Name & "'."
             End If
         End If
         Log "[Archival] GetOrCreateSheet: Created sheet '" & ws.Name & "'."
    End If

    ' --- Ensure sheet is visible ---
    If Not ws Is Nothing Then
         If ws.Visible <> xlSheetVisible Then ws.Visible = xlSheetVisible
    End If


    ' --- Apply title banner (A1:N1) ---
    Dim lastColLetter As String: lastColLetter = Module_Dashboard.DB_COL_COMMENTS ' Should be "N"
    On Error Resume Next ' Handle error if sheet is protected
    With ws.Range("A1:" & lastColLetter & "1") ' Use Constant
        If .MergeCells Then .UnMerge ' Ensure not already merged incorrectly
        .ClearContents
        .Merge
        .Value = sTitle
        .Font.Bold = True: .Font.Size = 16
        .Interior.Color = bannerColor
        .Font.Color = vbWhite
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .RowHeight = 32 ' Match main dashboard title height
    End With
    If Err.Number <> 0 Then
        Log "[Archival] GetOrCreateSheet: Warning - could not apply title formatting to '" & ws.Name & "'."
        Err.Clear
    End If
    On Error GoTo GetOrCreateSheet_Error ' Restore handler

    Set GetOrCreateSheet = ws ' Return the sheet object
    Exit Function ' Normal exit

GetOrCreateSheet_Error:
    Log "[Archival] ERROR in GetOrCreateSheet [" & Err.Number & "]: " & Err.Description
    ' Attempt to re-lock structure if we unlocked it and an error occurred
    If wbStructureWasLocked Then Module_Dashboard.ToggleWorkbookStructure (True)
    Set GetOrCreateSheet = Nothing ' Return Nothing on error

End Function

'--------------------------------------------------
'  Post-copy formatting + nav buttons + protection for Views
'--------------------------------------------------
Private Sub ApplyViewFormatting(ws As Worksheet, viewTag As String)
    Dim lastRow As Long
    Dim lastColLetter As String: lastColLetter = Module_Dashboard.DB_COL_COMMENTS ' Should be "N"
    Dim phaseColLetter As String: phaseColLetter = Module_Dashboard.DB_COL_PHASE ' Should be "L"
    Dim workflowColLetter As String: workflowColLetter = Module_Dashboard.DB_COL_WORKFLOW_LOCATION ' Should be "J"

    If ws Is Nothing Then Exit Sub

    Log "[Archival] ApplyViewFormatting: Applying formats to '" & ws.Name & "'."

    Application.ScreenUpdating = False ' Keep off during formatting
    On Error GoTo ApplyViewFormatting_Error ' Use specific handler

    ' --- Unprotect first ---
    ws.Unprotect Password:=Module_Dashboard.PW_WORKBOOK ' Use password if defined

    ' --- Apply Column Widths (Copy from Source Dashboard) ---
    Dim wsSrc As Worksheet
    On Error Resume Next
    Set wsSrc = ThisWorkbook.Worksheets(Module_Dashboard.DASHBOARD_SHEET_NAME)
    On Error GoTo ApplyViewFormatting_Error ' Restore handler

    If Not wsSrc Is Nothing Then
        wsSrc.Range("A1:" & lastColLetter & "1").Copy
        ws.Range("A1").PasteSpecial xlPasteColumnWidths
        Application.CutCopyMode = False
        Log "[Archival] ApplyViewFormatting: Copied column widths from main dashboard."
    Else
         Log "[Archival] ApplyViewFormatting: Warning - Could not copy column widths from main dashboard (not found)."
         ' Fallback: AutoFit A:N if source widths unavailable
         ws.Columns("A:" & lastColLetter).AutoFit
    End If

    ' --- Apply Number Formatting ---
    lastRow = ws.Cells(ws.rows.Count, "A").End(xlUp).Row
    ' Data starts at Row 4 on these views because we copy A1:N3 headers
    If lastRow >= 4 Then
        Log "[Archival] ApplyViewFormatting: Applying number formats to rows 4:" & lastRow
        ws.Range("D4:D" & lastRow).NumberFormat = "$#,##0.00"   ' Amount (D)
        ws.Range("E4:E" & lastRow).NumberFormat = "mm/dd/yyyy" ' Document Date (E)
        ws.Range("F4:F" & lastRow).NumberFormat = "mm/dd/yyyy" ' First Date Pulled (F)
        ws.Range("I4:I" & lastRow).NumberFormat = "0"           ' Pull Count (I)
        ' Add M (Last Contact Date) formatting if it's supposed to be a date
        On Error Resume Next ' Handle non-dates in column M
        ws.Range(Module_Dashboard.DB_COL_LASTCONTACT & "4:" & Module_Dashboard.DB_COL_LASTCONTACT & lastRow).NumberFormat = "mm/dd/yyyy" ' Last Contact (M)
        On Error GoTo ApplyViewFormatting_Error ' Restore handler
    End If

    ' --- Apply Conditional Formatting ---
    If lastRow >= 4 Then ' Data starts row 4
         Log "[Archival] ApplyViewFormatting: Applying conditional formatting..."
         ' Apply Phase formatting (Col L) - Uses the main module's helper
         Module_Dashboard.ApplyColorFormatting ws, 4 ' Start formatting data from row 4
         ' Apply Workflow formatting (Col J) - Uses the main module's helper
         Module_Dashboard.ApplyWorkflowLocationFormatting ws, 4 ' Start formatting data from row 4
         Log "[Archival] ApplyViewFormatting: Applied conditional formatting."
    End If


    ' --- Protection (Make read-only) ---
    ws.Cells.Locked = True ' Lock all cells on the view sheets

    ' --- Apply Freeze Panes ---
    ws.Activate ' Must activate to set freeze panes
    ActiveWindow.FreezePanes = False ' Unfreeze first
    ws.Range("A4").Select           ' Select cell below freeze rows (1-3)
    ActiveWindow.FreezePanes = True  ' Freeze Rows 1-3
    ws.Range("A1").Select ' Select A1 after freezing
    Log "[Archival] ApplyViewFormatting: Applied freeze panes."

     ' --- Final Protection ---
     ' Protects the sheet, allowing only selection of cells
    ws.Protect Password:=Module_Dashboard.PW_WORKBOOK, DrawingObjects:=True, Contents:=True, Scenarios:=True, _
        AllowFormattingCells:=False, AllowFormattingColumns:=False, AllowFormattingRows:=False, _
        AllowInsertingColumns:=False, AllowInsertingRows:=False, AllowInsertingHyperlinks:=False, _
        AllowDeletingColumns:=False, AllowDeletingRows:=False, AllowSorting:=False, AllowFiltering:=False, _
        AllowUsingPivotTables:=False
    Log "[Archival] ApplyViewFormatting: Applied sheet protection (Read-Only)."

    ' --- Add Navigation Buttons ---
    AddNavigationButtons ws ' Call helper to add buttons

ApplyViewFormatting_Cleanup:
    Application.ScreenUpdating = True
    Set wsSrc = Nothing
    Exit Sub ' Normal Exit

ApplyViewFormatting_Error:
    Log "[Archival] ERROR in ApplyViewFormatting [" & Err.Number & "]: " & Err.Description
    Resume ApplyViewFormatting_Cleanup ' Go to cleanup even on error

End Sub


'--------------------------------------------------
'  Navigation buttons – reuse ModernButton from Module_Dashboard
'--------------------------------------------------
Private Sub AddNavigationButtons(ws As Worksheet)
    If ws Is Nothing Then Exit Sub
    Log "[Archival] AddNavigationButtons: Adding buttons to '" & ws.Name & "'."

    Dim shp As Shape
    On Error Resume Next ' Ignore error if shape doesn't exist or sheet is protected

    ' --- Delete existing nav buttons on this sheet ---
     ws.Unprotect Password:=Module_Dashboard.PW_WORKBOOK ' Temporarily unprotect to delete shapes
    For Each shp In ws.Shapes
        ' Check prefix and if shape is anchored in row 2
        If Left$(shp.Name, 3) = "nav" And Not Intersect(shp.TopLeftCell, ws.rows(2)) Is Nothing Then
            shp.Delete
        End If
    Next shp
    If Err.Number <> 0 Then Log "[Archival] AddNavigationButtons: Warning - Error occurred during deletion of old buttons.": Err.Clear ' Clear error from loop/delete

    ' --- Add New Buttons (anchored G2, I2, K2) ---
    ' Ensure ModernButton is Public in Module_Dashboard
    Call Module_Dashboard.ModernButton(ws, "G2", "All Items", "modArchival.btnViewAll")
    Call Module_Dashboard.ModernButton(ws, "I2", "Active", "modArchival.btnViewActive")
    Call Module_Dashboard.ModernButton(ws, "K2", "Archive", "modArchival.btnViewArchive")

    ' --- Rename buttons for consistency ---
    ' Find by action rather than assuming order
    Dim btnAll As Shape, btnActive As Shape, btnArchive As Shape
    Set btnAll = Nothing: Set btnActive = Nothing: Set btnArchive = Nothing

    For Each shp In ws.Shapes ' Find the buttons we just added
       If shp.OnAction = "modArchival.btnViewAll" Then Set btnAll = shp
       If shp.OnAction = "modArchival.btnViewActive" Then Set btnActive = shp
       If shp.OnAction = "modArchival.btnViewArchive" Then Set btnArchive = shp
    Next shp

    If Not btnAll Is Nothing Then btnAll.Name = "navAll" Else Log "[Archival] AddNavButtons: All button not found."
    If Not btnActive Is Nothing Then btnActive.Name = "navActive" Else Log "[Archival] AddNavButtons: Active button not found."
    If Not btnArchive Is Nothing Then btnArchive.Name = "navArchive" Else Log "[Archival] AddNavButtons: Archive button not found."

    ' --- Re-protect sheet ---
    ws.Protect Password:=Module_Dashboard.PW_WORKBOOK, DrawingObjects:=True, Contents:=True, Scenarios:=True ' Re-apply basic protection needed for buttons

    If Err.Number <> 0 Then
        Log "[Archival] AddNavigationButtons: Error adding/naming/re-protecting buttons: " & Err.Description
        Err.Clear
    End If
    On Error GoTo 0 ' Restore default error handling

    Set shp = Nothing: Set btnAll = Nothing: Set btnActive = Nothing: Set btnArchive = Nothing
End Sub

'--------------------------------------------------
'  Tiny logger shim (delegates to real logger in Module_Dashboard)
'--------------------------------------------------
Private Sub Log(msg As String)
    On Error Resume Next ' Avoid error if logger sub doesn't exist or fails
    ' Call the logger in Module_Dashboard directly
    Module_Dashboard.LogUserEditsOperation msg
    If Err.Number <> 0 Then
        Debug.Print "Archival Log Error: Failed to call logger. Msg: " & msg ' Fallback to Immediate window
        Err.Clear
    End If
    On Error GoTo 0 ' Restore default error handling
End Sub

