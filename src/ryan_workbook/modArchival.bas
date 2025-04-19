'=====================================================================
'  Module :  modArchival  – Archive / Active / All dashboard views
'  Workbook: SQRCT
'  REVISED: Added Row Height fix, Blank Phase=Active fix, Naming Convention update
'           Includes ReDim Preserve fix & detailed logging.
'---------------------------------------------------------------------
Option Explicit

'--------------------------------------------------
' 1--  PHASE CATEGORISATION  (Based on user confirmation)
'--------------------------------------------------
Private Const ACTIVE_PHASES As String = "|First F/U|AF|RZ|KMH|RI|" ' Added leading/trailing pipes

'--------------------------------------------------
' 2--  SHEET NAMES & UI TEXT (NEW NAMING CONVENTION)
'--------------------------------------------------
' Using DASHBOARD_SHEET_NAME from Module_Dashboard via Public Const (e.g., "SQRCT_All")
Private Const SH_ACTIVE As String = "SQRCT_Active"     ' <<< RENAMED >>>
Private Const SH_ARCHIVE As String = "SQRCT_Archive"    ' <<< RENAMED >>>
Private Const TITLE_ACTIVE As String = _
        "STRATEGIC QUOTE RECOVERY & CONVERSION TRACKER – ACTIVE VIEW" ' Sheet title (Can remain same)
Private Const TITLE_ARCHIVE As String = _
        "STRATEGIC QUOTE RECOVERY & CONVERSION TRACKER – ARCHIVE VIEW" ' Sheet title (Can remain same)

'--------------------------------------------------
' 3--  PUBLIC BUTTON WRAPPERS  (Called by buttons)
'--------------------------------------------------
Public Sub btnViewActive()
    Module_Dashboard.DebugLog "btnViewActive", "ENTER"
    ' Refreshes the view THEN activates it
    Application.StatusBar = "Refreshing Active View..."
    RefreshAndActivate SH_ACTIVE ' Uses new constant
    Application.StatusBar = False
    Module_Dashboard.DebugLog "btnViewActive", "EXIT"
End Sub

Public Sub btnViewArchive()
    Module_Dashboard.DebugLog "btnViewArchive", "ENTER"
    ' Refreshes the view THEN activates it
    Application.StatusBar = "Refreshing Archive View..."
    RefreshAndActivate SH_ARCHIVE ' Uses new constant
    Application.StatusBar = False
    Module_Dashboard.DebugLog "btnViewArchive", "EXIT"
End Sub

Public Sub btnViewAll()
    Module_Dashboard.DebugLog "btnViewAll", "ENTER"
    ' Just activates the main dashboard
    Application.StatusBar = "Activating Main Dashboard..."
    On Error Resume Next
    ' Uses Public Const from Module_Dashboard, which should be updated to "SQRCT_All"
    ThisWorkbook.Worksheets(Module_Dashboard.DASHBOARD_SHEET_NAME).Activate
    If Err.Number <> 0 Then
         Module_Dashboard.DebugLog "btnViewAll", "ERROR activating sheet '" & Module_Dashboard.DASHBOARD_SHEET_NAME & "'. Err=" & Err.Number & ": " & Err.Description
         MsgBox "Could not activate sheet: " & Module_Dashboard.DASHBOARD_SHEET_NAME, vbExclamation
         Err.Clear
    Else
         Module_Dashboard.DebugLog "btnViewAll", "Activated sheet '" & Module_Dashboard.DASHBOARD_SHEET_NAME & "'"
    End If
    On Error GoTo 0
    Application.StatusBar = False
    Module_Dashboard.DebugLog "btnViewAll", "EXIT"
End Sub

'--------------------------------------------------
' 4--  TOP-LEVEL REFRESH ENTRY (Called by Module_Dashboard.RefreshDashboard)
' REVISED: Removed management of Application settings (Calc, Events, ScreenUpdate)
'--------------------------------------------------
Public Sub RefreshAllViews()
    Dim wsDash As Worksheet
    Dim originalStatusBar As Variant
    Dim t1 As Double: t1 = Timer ' Timer for this sub

    On Error GoTo ArchivalErrorHandler ' Use a specific handler for this sub

    Module_Dashboard.DebugLog "RefreshAllViews", "ENTER"

    ' --- Store & Set ONLY StatusBar ---
    originalStatusBar = Application.StatusBar ' Store current status bar text
    Application.StatusBar = "Refreshing Active/Archive views..."
    Module_Dashboard.DebugLog "RefreshAllViews", "StatusBar set."

    ' --- Get Main Dashboard Reference ---
    Module_Dashboard.DebugLog "RefreshAllViews", "Getting main dashboard sheet reference ('" & Module_Dashboard.DASHBOARD_SHEET_NAME & "')..." ' Uses updated constant name from Module_Dashboard
    Set wsDash = Nothing ' Initialize
    On Error Resume Next ' Temporarily ignore error getting sheet reference
    Set wsDash = ThisWorkbook.Worksheets(Module_Dashboard.DASHBOARD_SHEET_NAME) ' Uses updated constant name
    On Error GoTo ArchivalErrorHandler ' Restore main handler

    If wsDash Is Nothing Then
        Module_Dashboard.DebugLog "RefreshAllViews", "ERROR: Main dashboard sheet '" & Module_Dashboard.DASHBOARD_SHEET_NAME & "' not found. Cannot refresh views."
        GoTo ArchivalCleanup ' Go to cleanup
    End If
    Module_Dashboard.DebugLog "RefreshAllViews", "Got main dashboard sheet: '" & wsDash.Name & "'"

    ' --- Refresh Views ---
    Module_Dashboard.DebugLog "RefreshAllViews", "Calling RefreshActiveView..."
    RefreshActiveView wsDash   ' Pass source sheet
    Module_Dashboard.DebugLog "RefreshAllViews", "Returned from RefreshActiveView."

    Module_Dashboard.DebugLog "RefreshAllViews", "Calling RefreshArchiveView..."
    RefreshArchiveView wsDash  ' Pass source sheet
    Module_Dashboard.DebugLog "RefreshAllViews", "Returned from RefreshArchiveView."

    Module_Dashboard.DebugLog "RefreshAllViews", "Calling AddNavigationButtons for main dashboard ('" & wsDash.Name & "')..."
    AddNavigationButtons wsDash ' Add buttons to main dashboard
    Module_Dashboard.DebugLog "RefreshAllViews", "Returned from AddNavigationButtons for main dashboard."

ArchivalCleanup:    ' Cleanup label for both normal exit and error
    Module_Dashboard.DebugLog "RefreshAllViews", "Cleanup Label Reached..."
    On Error Resume Next ' Ignore errors during cleanup itself

    ' --- Restore ONLY StatusBar ---
    Application.StatusBar = originalStatusBar ' Restore original status bar text
    Set wsDash = Nothing

    Module_Dashboard.DebugLog "RefreshAllViews", "EXIT (Cleanup Complete). Time: " & Format(Timer - t1, "0.00") & "s"
    Exit Sub ' Explicit Exit for clarity after normal completion or cleanup jump

ArchivalErrorHandler: ' Error handler for this sub
    Module_Dashboard.DebugLog "RefreshAllViews", "!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!"
    Module_Dashboard.DebugLog "RefreshAllViews", "ERROR Handler! Err=" & Err.Number & ": " & Err.Description & " near line " & Erl
    Module_Dashboard.DebugLog "RefreshAllViews", "!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!"
    MsgBox "An error occurred while refreshing the Active/Archive views:" & vbCrLf & Err.Description, vbCritical, "Archival Refresh Error"
    Resume ArchivalCleanup ' Go to cleanup after logging error

End Sub

'*********************** PRIVATE SECTION  ***************************

'--------------------------------------------------
'  Refresh & activate a single view (helper)
'  REVISED: Removed management of Application settings
'--------------------------------------------------
Private Sub RefreshAndActivate(viewName As String) ' viewName will be SQRCT_Active or SQRCT_Archive
    Dim wsDash As Worksheet
    Dim originalStatusBar As Variant
    Dim t1 As Double: t1 = Timer

    On Error GoTo RefreshActivateErrorHandler ' Specific handler

    Module_Dashboard.DebugLog "RefreshAndActivate", "ENTER for viewName='" & viewName & "'"

    ' --- Store & Set ONLY StatusBar ---
    originalStatusBar = Application.StatusBar
    Application.StatusBar = "Refreshing " & viewName & "..." ' Update status
    Module_Dashboard.DebugLog "RefreshAndActivate", "StatusBar set."

    ' --- Get Main Dashboard Reference ---
    Module_Dashboard.DebugLog "RefreshAndActivate", "Getting main dashboard sheet reference ('" & Module_Dashboard.DASHBOARD_SHEET_NAME & "')..."
    Set wsDash = Nothing
    On Error Resume Next
    Set wsDash = ThisWorkbook.Worksheets(Module_Dashboard.DASHBOARD_SHEET_NAME) ' Uses updated constant
    On Error GoTo RefreshActivateErrorHandler ' Restore handler

    If wsDash Is Nothing Then
        Module_Dashboard.DebugLog "RefreshAndActivate", "ERROR in RefreshAndActivate: Main dashboard sheet not found."
        GoTo RefreshActivateCleanup ' Go to cleanup
    End If
    Module_Dashboard.DebugLog "RefreshAndActivate", "Got main dashboard sheet: '" & wsDash.Name & "'"

    ' --- Refresh the specific view ---
    Module_Dashboard.DebugLog "RefreshAndActivate", "Selecting view to refresh based on viewName='" & viewName & "'"
    Select Case viewName
        Case SH_ACTIVE ' Uses updated constant
            Module_Dashboard.DebugLog "RefreshAndActivate", "Calling RefreshActiveView..."
            RefreshActiveView wsDash
            Module_Dashboard.DebugLog "RefreshAndActivate", "Returned from RefreshActiveView."
        Case SH_ARCHIVE ' Uses updated constant
            Module_Dashboard.DebugLog "RefreshAndActivate", "Calling RefreshArchiveView..."
            RefreshArchiveView wsDash
            Module_Dashboard.DebugLog "RefreshAndActivate", "Returned from RefreshArchiveView."
        Case Else
            Module_Dashboard.DebugLog "RefreshAndActivate", "ERROR: Invalid view name '" & viewName & "'."
            GoTo RefreshActivateCleanup
    End Select

    ' --- Activate the sheet ---
    Module_Dashboard.DebugLog "RefreshAndActivate", "Attempting to activate sheet '" & viewName & "'..."
    On Error Resume Next ' Ignore error activating (sheet might be hidden temporarily)
    ThisWorkbook.Worksheets(viewName).Activate ' Uses viewName directly (e.g. "SQRCT_Active")
    If Err.Number <> 0 Then
        Module_Dashboard.DebugLog "RefreshAndActivate", "ERROR: Could not activate sheet '" & viewName & "'. Error: " & Err.Description
        Err.Clear ' Clear the error
    Else
        Module_Dashboard.DebugLog "RefreshAndActivate", "Activated sheet '" & viewName & "'"
    End If
    On Error GoTo RefreshActivateErrorHandler ' Restore handler

RefreshActivateCleanup:
    Module_Dashboard.DebugLog "RefreshAndActivate", "Cleanup Label Reached for '" & viewName & "'"
    On Error Resume Next ' Ignore errors during cleanup
    ' --- Restore ONLY StatusBar ---
    Application.StatusBar = originalStatusBar ' Clear status bar
    Set wsDash = Nothing
    Module_Dashboard.DebugLog "RefreshAndActivate", "EXIT (Cleanup Complete) for '" & viewName & "'. Time: " & Format(Timer - t1, "0.00") & "s"
    Exit Sub

RefreshActivateErrorHandler:
     Module_Dashboard.DebugLog "RefreshAndActivate", "!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!"
     Module_Dashboard.DebugLog "RefreshAndActivate", "ERROR Handler! viewName='" & viewName & "'. Err=" & Err.Number & ": " & Err.Description & " near line " & Erl
     Module_Dashboard.DebugLog "RefreshAndActivate", "!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!"
     MsgBox "An error occurred while refreshing/activating the " & viewName & " view:" & vbCrLf & Err.Description, vbExclamation, "View Activation Error"
     Resume RefreshActivateCleanup ' Go to cleanup

End Sub

'--------------------------------------------------
'  ACTIVE  VIEW  (filter-in rows *where phase is active*)
'--------------------------------------------------
Private Sub RefreshActiveView(wsSrc As Worksheet) ' Takes source ws as parameter
    Dim wsTgt As Worksheet
    Dim t1 As Double: t1 = Timer

    On Error GoTo RefreshActive_Error ' Add error handling for this sub
    Module_Dashboard.DebugLog "RefreshActiveView", "ENTER. Source sheet='" & wsSrc.Name & "'"

    Module_Dashboard.DebugLog "RefreshActiveView", "Calling GetOrCreateSheet for '" & SH_ACTIVE & "'..." ' Uses updated constant
    Set wsTgt = GetOrCreateSheet(SH_ACTIVE, TITLE_ACTIVE, RGB(0, 110, 0)) ' Dark Green Banner

    If wsTgt Is Nothing Then
        Module_Dashboard.DebugLog "RefreshActiveView", "GetOrCreateSheet failed. Aborting."
        GoTo RefreshActive_Cleanup ' Go to cleanup
    End If
    Module_Dashboard.DebugLog "RefreshActiveView", "Got target sheet: '" & wsTgt.Name & "'"

    Module_Dashboard.DebugLog "RefreshActiveView", "Calling CopyFilteredRows (keepArchived=False)..."
    CopyFilteredRows wsSrc, wsTgt, keepArchived:=False ' Uses updated IsPhaseArchived logic internally
    Module_Dashboard.DebugLog "RefreshActiveView", "Returned from CopyFilteredRows."

    ' --- Add Defensive Check before Formatting ---
    Module_Dashboard.DebugLog "RefreshActiveView", "Performing defensive checks before formatting..."
    If wsTgt Is Nothing Then ' Double check object variable just in case
         Module_Dashboard.DebugLog "RefreshActiveView", "Defensive Check FAIL: wsTgt object became Nothing unexpectedly."
    ElseIf Not SheetExists(wsTgt.Name) Then ' Check if sheet still exists by name
         Module_Dashboard.DebugLog "RefreshActiveView", "Defensive Check FAIL: Target sheet '" & SH_ACTIVE & "' no longer exists before formatting." ' Uses updated constant
    Else
         ' Sheet exists, proceed with formatting
         Module_Dashboard.DebugLog "RefreshActiveView", "Defensive Check PASS: Target sheet valid, calling ApplyViewFormatting..."
         ApplyViewFormatting wsTgt, "Active" ' Apply formatting and buttons
         Module_Dashboard.DebugLog "RefreshActiveView", "Returned from ApplyViewFormatting."
    End If
    ' --- End Defensive Check ---

    Module_Dashboard.DebugLog "RefreshActiveView", "Process completed normally (before cleanup)."

RefreshActive_Cleanup: ' Cleanup label
    Module_Dashboard.DebugLog "RefreshActiveView", "Cleanup Label Reached."
    On Error Resume Next ' Ignore errors during cleanup
    ' --- Re-enable Events ---
    If Not Application.EnableEvents Then
        Module_Dashboard.DebugLog "RefreshActiveView", "Cleanup: Re-enabling Application.EnableEvents."
        Application.EnableEvents = True
    Else
        Module_Dashboard.DebugLog "RefreshActiveView", "Cleanup: Application.EnableEvents already True."
    End If
    Set wsTgt = Nothing
    Module_Dashboard.DebugLog "RefreshActiveView", "EXIT (Cleanup Complete). Time: " & Format(Timer - t1, "0.00") & "s"
    Exit Sub ' Normal exit from cleanup

RefreshActive_Error:
    Module_Dashboard.DebugLog "RefreshActiveView", "!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!"
    Module_Dashboard.DebugLog "RefreshActiveView", "ERROR Handler! Err=" & Err.Number & ": " & Err.Description & " near line " & Erl
    Module_Dashboard.DebugLog "RefreshActiveView", "!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!"
    Resume RefreshActive_Cleanup ' Go to cleanup on error

End Sub


'--------------------------------------------------
'  ARCHIVE VIEW  (filter-in rows *where phase is archived*)
'--------------------------------------------------
Private Sub RefreshArchiveView(wsSrc As Worksheet) ' Takes source ws as parameter
    Dim wsTgt As Worksheet
    Dim t1 As Double: t1 = Timer

    On Error GoTo RefreshArchive_Error ' Add error handling for this sub
    Module_Dashboard.DebugLog "RefreshArchiveView", "ENTER. Source sheet='" & wsSrc.Name & "'"

    Module_Dashboard.DebugLog "RefreshArchiveView", "Calling GetOrCreateSheet for '" & SH_ARCHIVE & "'..." ' Uses updated constant
    Set wsTgt = GetOrCreateSheet(SH_ARCHIVE, TITLE_ARCHIVE, RGB(150, 40, 40)) ' Dark Red Banner

    If wsTgt Is Nothing Then
        Module_Dashboard.DebugLog "RefreshArchiveView", "GetOrCreateSheet failed. Aborting."
        GoTo RefreshArchive_Cleanup ' Go to cleanup
    End If
    Module_Dashboard.DebugLog "RefreshArchiveView", "Got target sheet: '" & wsTgt.Name & "'"

    Module_Dashboard.DebugLog "RefreshArchiveView", "Calling CopyFilteredRows (keepArchived=True)..." ' Uses updated IsPhaseArchived logic internally
    CopyFilteredRows wsSrc, wsTgt, keepArchived:=True
    Module_Dashboard.DebugLog "RefreshArchiveView", "Returned from CopyFilteredRows."

    Module_Dashboard.DebugLog "RefreshArchiveView", "Calling ApplyViewFormatting..."
    ApplyViewFormatting wsTgt, "Archive" ' Apply formatting and buttons
    Module_Dashboard.DebugLog "RefreshArchiveView", "Returned from ApplyViewFormatting."

RefreshArchive_Cleanup: ' Cleanup label
    Module_Dashboard.DebugLog "RefreshArchiveView", "Cleanup Label Reached."
    On Error Resume Next ' Ignore errors during cleanup
    Set wsTgt = Nothing
    Module_Dashboard.DebugLog "RefreshArchiveView", "EXIT (Cleanup Complete). Time: " & Format(Timer - t1, "0.00") & "s"
    Exit Sub ' Normal exit from cleanup

RefreshArchive_Error:
    Module_Dashboard.DebugLog "RefreshArchiveView", "!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!"
    Module_Dashboard.DebugLog "RefreshArchiveView", "ERROR Handler! Err=" & Err.Number & ": " & Err.Description & " near line " & Erl
    Module_Dashboard.DebugLog "RefreshArchiveView", "!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!"
    Resume RefreshArchive_Cleanup ' Go to cleanup on error

End Sub

'--------------------------------------------------
'  Copy rows from main dashboard based on phase test (Uses Array)
'  REVISED: Includes enhanced logging for filter loop.
'           Avoids ReDim Preserve issue by using second array copy.
'           Includes explicit Unprotect and specific Write Error trap.
'--------------------------------------------------
Private Sub CopyFilteredRows(wsSrc As Worksheet, wsTgt As Worksheet, _
                             keepArchived As Boolean)
    Dim lastSrcRow As Long, destRow As Long, r As Long
    Dim phaseValue As String
    Dim arrData As Variant      ' Array to hold source data A:N
    Dim i As Long, j As Long    ' Loop counters
    Dim outputData() As Variant ' Array to hold rows TO BE copied (potentially oversized)
    Dim finalOutputData() As Variant ' Array sized exactly for output
    Dim outputRowCount As Long
    Dim srcRange As Range
    Dim phaseColIndex As Long
    Dim numCols As Long
    Dim t1 As Double: t1 = Timer

    Module_Dashboard.DebugLog "CopyFilteredRows", "ENTER. Source='" & wsSrc.Name & "', Target='" & wsTgt.Name & "', KeepArchived=" & keepArchived

    ' --- Define Source Data Range (A to N) ---
    Dim lastColLetter As String: lastColLetter = Module_Dashboard.DB_COL_COMMENTS ' Should be "N"
    Dim firstDataRow As Long: firstDataRow = 4 ' Data starts at row 4 on main dash
    Dim expectedCols As Long: expectedCols = 14 ' Expected columns A-N

    On Error GoTo CopyFilteredRows_Error ' Add local error handler

    ' --- Ensure Target Sheet is Writable ---
    Module_Dashboard.DebugLog "CopyFilteredRows", "Ensuring target sheet '" & wsTgt.Name & "' is writable..."
    On Error Resume Next ' Handle error if already unprotected or password wrong
    If wsTgt.ProtectContents Then
        Module_Dashboard.DebugLog "CopyFilteredRows", "Target sheet is protected. Attempting unprotect..."
        wsTgt.Unprotect Password:=Module_Dashboard.PW_WORKBOOK ' Use constant
        If Err.Number <> 0 Then
             Module_Dashboard.DebugLog "CopyFilteredRows", "WARNING: Failed to unprotect target sheet. Write may fail. Err=" & Err.Number & ": " & Err.Description
             Err.Clear
        Else
             Module_Dashboard.DebugLog "CopyFilteredRows", "Target sheet successfully unprotected."
        End If
    Else
         Module_Dashboard.DebugLog "CopyFilteredRows", "Target sheet is already unprotected."
    End If
    On Error GoTo CopyFilteredRows_Error ' Restore main error handler for this sub
    ' --- End Ensure Target Sheet is Writable ---

    lastSrcRow = wsSrc.Cells(wsSrc.rows.Count, "A").End(xlUp).Row
    Module_Dashboard.DebugLog "CopyFilteredRows", "Source last row = " & lastSrcRow

    ' --- Handle No Data Case ---
    If lastSrcRow < firstDataRow Then
        Module_Dashboard.DebugLog "CopyFilteredRows", "No data rows found on source sheet. Copying headers only."
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
        On Error GoTo CopyFilteredRows_Error ' Restore Handler
        Module_Dashboard.DebugLog "CopyFilteredRows", "EXIT (No data case). Time: " & Format(Timer - t1, "0.00") & "s"
        Exit Sub
    End If

    ' --- Read source data into array for faster processing ---
    Module_Dashboard.DebugLog "CopyFilteredRows", "Reading source range A" & firstDataRow & ":" & lastColLetter & lastSrcRow & " into array..."
    On Error Resume Next
    Set srcRange = wsSrc.Range("A" & firstDataRow & ":" & lastColLetter & lastSrcRow) ' A4:N<lastRow>
    arrData = srcRange.Value2 ' This line reads the data
    If Err.Number <> 0 Then
         Module_Dashboard.DebugLog "CopyFilteredRows", "ERROR reading range into arrData. Err=" & Err.Number & ": " & Err.Description
         Err.Clear ' Clear error before checking IsArray
    End If
    On Error GoTo CopyFilteredRows_Error ' Restore Handler

    ' Check if reading failed or didn't produce an array
    If Not IsArray(arrData) Then
        Module_Dashboard.DebugLog "CopyFilteredRows", "ERROR: arrData is not an array after reading range. Exiting."
        Exit Sub ' Or handle error appropriately
    End If
    Module_Dashboard.DebugLog "CopyFilteredRows", "Read successful. UBound(arrData, 1)=" & UBound(arrData, 1)

    ' >>> Array Dimension Handling (Check actual columns read) <<<
    Module_Dashboard.DebugLog "CopyFilteredRows", "Checking actual array dimensions read..."
    Dim actualCols As Long
    On Error Resume Next ' Handle error getting UBound if arrData is invalid
    actualCols = UBound(arrData, 2)
    If Err.Number <> 0 Then
        Module_Dashboard.DebugLog "CopyFilteredRows", "ERROR getting UBound(arrData, 2). Setting numCols=0. Err=" & Err.Number & ": " & Err.Description
        numCols = 0 ' Handle error getting UBound
        Err.Clear
    Else
        numCols = actualCols ' Use the actual column count read
    End If
    On Error GoTo CopyFilteredRows_Error ' Restore Handler
    Module_Dashboard.DebugLog "CopyFilteredRows", "Actual columns read = " & actualCols & ". numCols set to " & numCols

    ' --- Handle Single Row Case (Value2 read gives 1D array) ---
    If lastSrcRow = firstDataRow Then
         Dim is2D As Boolean
         On Error Resume Next
         Dim dummy As Long: dummy = UBound(arrData, 2) ' Check if 2nd dimension exists
         is2D = (Err.Number = 0)
         On Error GoTo CopyFilteredRows_Error ' Restore Handler

         If Not is2D Then ' It's 1D
             Module_Dashboard.DebugLog "CopyFilteredRows", "Handling single row case (1D array found)."
             Dim tempArr() As Variant
             Dim itemsIn1D As Long: itemsIn1D = UBound(arrData, 1) - LBound(arrData, 1) + 1
             numCols = itemsIn1D ' Number of columns is number of items in 1D array
             Module_Dashboard.DebugLog "CopyFilteredRows", "1D array has " & itemsIn1D & " items. Setting numCols=" & numCols
             ReDim tempArr(1 To 1, 1 To numCols) ' Size based on actual items read

             For j = LBound(arrData, 1) To UBound(arrData, 1)
                  tempArr(1, j) = arrData(j)
             Next j
             Erase arrData
             arrData = tempArr ' Replace original 1D array with 2D
             Module_Dashboard.DebugLog "CopyFilteredRows", "Converted 1D to 2D array (1x" & numCols & ")."
         End If
    End If

    ' --- Prepare initial (potentially oversized) output array ---
    If numCols = 0 Then ' Safety check if numCols failed to be set
        Module_Dashboard.DebugLog "CopyFilteredRows", "ERROR: numCols is 0 before preparing output array. Exiting."
        Exit Sub
    End If
    Module_Dashboard.DebugLog "CopyFilteredRows", "Preparing initial output array: 1 To " & UBound(arrData, 1) & ", 1 To " & numCols
    ReDim outputData(1 To UBound(arrData, 1), 1 To numCols) ' Max possible size based on source rows
    outputRowCount = 0

    ' --- Get the 1-based index for the Phase column (L) within the array ---
    On Error Resume Next ' In case constant refers to invalid column
    phaseColIndex = wsSrc.Columns(Module_Dashboard.DB_COL_PHASE).Column
    If Err.Number <> 0 Or phaseColIndex = 0 Or phaseColIndex > numCols Then
        Module_Dashboard.DebugLog "CopyFilteredRows", "ERROR - Invalid Phase Column '" & Module_Dashboard.DB_COL_PHASE & "' index (" & phaseColIndex & "). numCols=" & numCols
        Err.Clear ' Clear error before exiting
        Erase arrData ' Clean up input array
        Exit Sub
    End If
    On Error GoTo CopyFilteredRows_Error ' Restore Handler
    Module_Dashboard.DebugLog "CopyFilteredRows", "Phase column index (L) = " & phaseColIndex

    ' --- START OF REPLACEMENT BLOCK ---
    ' --- Filter rows in memory into potentially oversized outputData ---
    Module_Dashboard.DebugLog "CopyFilteredRows", _
            "Filtering " & UBound(arrData, 1) & " rows …  keepArchived=" & _
            keepArchived & ", phaseColIndex=" & phaseColIndex

    outputRowCount = 0
    For r = LBound(arrData, 1) To UBound(arrData, 1)
        phaseValue = Trim$(CStr(arrData(r, phaseColIndex)))

        ' sample the first 10 rows so we can see what the code sees
        If r <= 10 Then _
            Module_Dashboard.DebugLog "FilterSample", _
            "r=" & r & ", phase='" & phaseValue & _
            "', IsPhaseArchived=" & IsPhaseArchived(phaseValue)

        If IsPhaseArchived(phaseValue) = keepArchived Then
            outputRowCount = outputRowCount + 1
            For i = 1 To numCols
                outputData(outputRowCount, i) = arrData(r, i)
            Next i
        End If
    Next r

    Module_Dashboard.DebugLog "CopyFilteredRows", _
            "Filtering complete.  Matching rows = " & outputRowCount
    ' *** sanity-check: if nothing was copied but the source had data, tell us ***
    If outputRowCount = 0 And UBound(arrData, 1) >= 1 Then
        Module_Dashboard.DebugLog "CopyFilteredRows", _
            "WARNING – zero rows met the criteria; source had " & _
            UBound(arrData, 1) & " rows."
    End If
    ' --- END OF REPLACEMENT BLOCK ---

    ' --- Write results to Target Sheet ---
    Module_Dashboard.DebugLog "CopyFilteredRows", "Writing results to target sheet '" & wsTgt.Name & "'..."
    wsTgt.Cells.ClearContents                       ' Clear target completely first
    Module_Dashboard.DebugLog "CopyFilteredRows", "Cleared target sheet."
    wsSrc.Range("A1:" & lastColLetter & "3").Copy wsTgt.Range("A1") ' Copy A1:N3 (Title/Control/Headers)
    Module_Dashboard.DebugLog "CopyFilteredRows", "Copied headers A1:" & lastColLetter & "3."

    '----- WRITE THE FILTERED ROWS -----
    If outputRowCount > 0 Then

        ' --- FIX for ReDim Preserve 32-bit Issue: Copy to correctly sized array ---
        Module_Dashboard.DebugLog "CopyFilteredRows", "Preparing finalOutputData array sized " & outputRowCount & "x" & numCols
        ReDim finalOutputData(1 To outputRowCount, 1 To numCols) ' Size exactly right
        Module_Dashboard.DebugLog "CopyFilteredRows", "Copying " & outputRowCount & " rows from temp outputData to finalOutputData..."
        For r = 1 To outputRowCount
            For i = 1 To numCols
                finalOutputData(r, i) = outputData(r, i) ' Copy relevant rows/cols
            Next i
        Next r
        Module_Dashboard.DebugLog "CopyFilteredRows", "Finished copying to finalOutputData."
        Erase outputData ' Release memory of oversized array
        ' --- End Fix ---

        ' --- Pre-Write Checks ---
        Dim isProtected As Boolean
        Dim isMerged As Boolean
        On Error Resume Next ' Check properties safely
        isProtected = wsTgt.ProtectContents
        isMerged = wsTgt.Range("A4").MergeCells ' Check A4 specifically for merge
        On Error GoTo CopyFilteredRows_Error ' Restore main handler
        Module_Dashboard.DebugLog "CopyFilteredRows", "PRE-WRITE CHECK: wsTgt.ProtectContents = " & isProtected
        Module_Dashboard.DebugLog "CopyFilteredRows", "PRE-WRITE CHECK: wsTgt.Range(""A4"").MergeCells = " & isMerged
        ' --- End Pre-Write Checks ---

        Module_Dashboard.DebugLog "CopyFilteredRows", "Attempting write finalOutputData to " & wsTgt.Name & ".Range(""A4"")..."
        On Error GoTo WriteErr ' *** Trap errors JUST for the write ***
        wsTgt.Range("A4").Resize(outputRowCount, numCols).Value = finalOutputData ' <<< WRITE finalOutputData >>>
        On Error GoTo CopyFilteredRows_Error ' *** Restore main handler AFTER successful write ***

        ' If we get here, write was successful
        Module_Dashboard.DebugLog "CopyFilteredRows", "Wrote " & outputRowCount & " rows to '" & wsTgt.Name & "'."
        GoTo AfterWrite ' Skip the error handler block

WriteErr: ' *** Specific handler for write error ***
        Module_Dashboard.DebugLog "CopyFilteredRows", "*** WRITE FAILED *** Err " & Err.Number & ": " & Err.Description
        MsgBox "CopyFilteredRows failed on '" & wsTgt.Name & "'" & vbCrLf & _
               "Err " & Err.Number & ": " & Err.Description, vbCritical, "Data Write Error"
        On Error GoTo CopyFilteredRows_Error ' Restore main handler before exiting
        Erase arrData: If IsArray(finalOutputData) Then Erase finalOutputData ' Clean up arrays
        Exit Sub ' Exit after write failure

AfterWrite: ' Label to jump to after successful write
        ' Any code needed after successful write could go here

    Else ' outputRowCount = 0
        Module_Dashboard.DebugLog "CopyFilteredRows", "No rows matched filter criteria for '" & wsTgt.Name & "'. Clearing data area."
        ' Ensure data area is clear if no rows written
        wsTgt.Range("A4:" & lastColLetter & wsTgt.rows.Count).ClearContents
    End If
    ' --- End of write block ---


    ' Add count below timestamp area (handle potential merge)
    Module_Dashboard.DebugLog "CopyFilteredRows", "Adding row count to G2:J2 area..."
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
    On Error GoTo CopyFilteredRows_Error ' Restore Handler

    ' Clean up memory
    Erase arrData
    If IsArray(finalOutputData) Then Erase finalOutputData ' Erase the correct array
    Module_Dashboard.DebugLog "CopyFilteredRows", "EXIT (Normal). Time: " & Format(Timer - t1, "0.00") & "s"
    Exit Sub ' Normal Exit

CopyFilteredRows_Error:
    If Err.Number <> 0 Then ' <<< ADDED CHECK TO ONLY LOG REAL ERRORS >>>
        Module_Dashboard.DebugLog "CopyFilteredRows", "!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!"
        Module_Dashboard.DebugLog "CopyFilteredRows", "ERROR Handler! Err=" & Err.Number & ": " & Err.Description & " near line " & Erl
        Module_Dashboard.DebugLog "CopyFilteredRows", "!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!"
    End If
    ' Clean up memory even on error
    If IsArray(arrData) Then Erase arrData
    If IsArray(outputData) Then Erase outputData ' Original temp array might still exist if error was early
    If IsArray(finalOutputData) Then Erase finalOutputData ' Final array
    ' Let error propagate back to the caller's error handler
End Sub


'--------------------------------------------------
'  Phase test helper - Checks if phase is considered Active
'--------------------------------------------------
'--------------------------------------------------
' TRUE if the phase counts as active
'--------------------------------------------------
Private Function IsPhaseActive(ph As String) As Boolean
    ph = UCase$(Trim$(ph))

    ' ? blanks are ACTIVE  ?
    If Len(ph) = 0 Then
        IsPhaseActive = True
        Exit Function
    End If

    ' pipe-delimited whitelist
    IsPhaseActive = (InStr(1, ACTIVE_PHASES, "|" & ph & "|") > 0)
End Function


'--------------------------------------------------
'  Phase test helper - checks if a phase is Archived
'  REVISED: Treats blanks as ACTIVE, uses explicit list for Archived phases.
'--------------------------------------------------
Private Function IsPhaseArchived(ByVal phase As String) As Boolean
    ' A phase is considered archived ONLY if it's explicitly in the list below.
    ' Blanks and any other non-listed phases are considered ACTIVE.

    ' --- Trim spaces and handle blanks first ---
    phase = Trim$(phase) ' Remove leading/trailing spaces
    If Len(phase) = 0 Then
        IsPhaseArchived = False ' Blank phase = ACTIVE (NOT Archived)
        Exit Function
    End If
    ' --- End blank check ---

    ' --- Explicitly check if phase is in the ARCHIVED list ---
    Select Case UCase$(phase) ' UCase makes comparison case-insensitive
        ' ***** ADJUST THIS LIST TO MATCH YOUR ARCHIVED PHASES *****
        Case "CONVERTED", "DECLINED", "CLOSED (EXTRA ORDER)", "CLOSED", "TEXAS (NO F/U)", "NO RESPONSE", "WW/OM", "LONG-TERM F/U"
            IsPhaseArchived = True ' Found in the Archived list
        Case Else
            IsPhaseArchived = False ' Any other non-blank phase = ACTIVE (NOT Archived)
    End Select
    ' ***** END OF LIST TO ADJUST *****

End Function

'--------------------------------------------------
'  SheetExists - Helper to check if a sheet name exists
'--------------------------------------------------
Private Function SheetExists(sName As String) As Boolean
    ' Returns True if a sheet with the given name exists in ThisWorkbook
    Dim sh As Object ' Use Object to check any sheet type
    On Error Resume Next ' Prevent error if sheet doesn't exist
    Set sh = ThisWorkbook.Sheets(sName)
    SheetExists = Not sh Is Nothing
    On Error GoTo 0 ' Restore default error handling
    Set sh = Nothing
End Function

'--------------------------------------------------
'  Create sheet if missing + drop coloured title bar
'--------------------------------------------------
Private Function GetOrCreateSheet(sName As String, sTitle As String, _
                                  bannerColor As Long) As Worksheet
    Dim ws As Worksheet
    Dim wbStructureWasLocked As Boolean
    Dim t1 As Double: t1 = Timer

    Module_Dashboard.DebugLog "GetOrCreateSheet", "ENTER. sName='" & sName & "'"
    On Error GoTo GetOrCreateSheet_Error ' Use specific handler

    ' --- Check if sheet exists ---
    Module_Dashboard.DebugLog "GetOrCreateSheet", "Checking if sheet exists..."
    Set ws = Nothing ' Ensure ws is Nothing initially
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(sName) ' Try to get as Worksheet specifically
    If Err.Number <> 0 Then Err.Clear ' Clear error if not found or wrong type
    On Error GoTo GetOrCreateSheet_Error ' Restore handler

    If ws Is Nothing Then
        ' Sheet does not exist as a worksheet, or doesn't exist at all.
        ' Check if *any* sheet type exists with this name (e.g., Chart)
        Dim anySheet As Object
        Set anySheet = Nothing
        On Error Resume Next
        Set anySheet = ThisWorkbook.Sheets(sName)
        On Error GoTo GetOrCreateSheet_Error
        If Not anySheet Is Nothing Then
             Module_Dashboard.DebugLog "GetOrCreateSheet", "WARNING: A non-worksheet object named '" & sName & "' exists (Type: " & TypeName(anySheet) & "). Renaming it."
             On Error Resume Next ' Handle error renaming
             anySheet.Name = sName & "_Old_" & Format(Now, "hhmmss")
             If Err.Number <> 0 Then
                  Module_Dashboard.DebugLog "GetOrCreateSheet", "ERROR: Failed to rename existing non-worksheet '" & sName & "'. Aborting. Err=" & Err.Number & ": " & Err.Description
                  Set GetOrCreateSheet = Nothing: Exit Function
             End If
             Err.Clear
             On Error GoTo GetOrCreateSheet_Error
             Set anySheet = Nothing ' Release object
        End If

        ' Proceed with creation
        Module_Dashboard.DebugLog "GetOrCreateSheet", "Sheet '" & sName & "' not found or was renamed. Attempting creation..."
        ' --- Need to temporarily unprotect workbook structure to add sheet ---
        wbStructureWasLocked = ThisWorkbook.ProtectStructure
        Module_Dashboard.DebugLog "GetOrCreateSheet", "Structure locked = " & wbStructureWasLocked
        If wbStructureWasLocked Then
             Module_Dashboard.DebugLog "GetOrCreateSheet", "Attempting structure unlock..."
            If Not Module_Dashboard.ToggleWorkbookStructure(False) Then ' Use Toggle function from Module_Dashboard
                Module_Dashboard.DebugLog "GetOrCreateSheet", "FATAL - Failed to unprotect workbook structure to add sheet '" & sName & "'."
                Set GetOrCreateSheet = Nothing ' Return Nothing
                Exit Function ' Cannot proceed
            End If
             Module_Dashboard.DebugLog "GetOrCreateSheet", "Structure unlocked."
        End If

        ' --- Add the sheet ---
        Module_Dashboard.DebugLog "GetOrCreateSheet", "Adding new worksheet..."
        On Error Resume Next ' Handle error during Add
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        On Error GoTo GetOrCreateSheet_Error ' Restore handler
        If ws Is Nothing Then ' Check if Add failed
             Module_Dashboard.DebugLog "GetOrCreateSheet", "FATAL - Failed to add new sheet (Set ws returned Nothing)."
             ' Attempt to re-lock structure if we unlocked it
             If wbStructureWasLocked Then Module_Dashboard.ToggleWorkbookStructure (True)
             Set GetOrCreateSheet = Nothing
             Exit Function
        End If
        Module_Dashboard.DebugLog "GetOrCreateSheet", "Added new sheet object: '" & ws.Name & "'"

        ' --- Name the sheet ---
        Module_Dashboard.DebugLog "GetOrCreateSheet", "Attempting to name new sheet '" & sName & "'..."
        On Error Resume Next ' Handle error during Name
        ws.Name = sName
        If Err.Number <> 0 Then
             Module_Dashboard.DebugLog "GetOrCreateSheet", "WARNING - Failed to name sheet '" & sName & "'. Using default '" & ws.Name & "'. Error: " & Err.Description
             Err.Clear
        End If
        On Error GoTo GetOrCreateSheet_Error ' Restore handler
        Module_Dashboard.DebugLog "GetOrCreateSheet", "Sheet name is now '" & ws.Name & "'"

         ' --- Re-lock structure if needed ---
         If wbStructureWasLocked Then
             Module_Dashboard.DebugLog "GetOrCreateSheet", "Attempting structure re-lock..."
             If Not Module_Dashboard.ToggleWorkbookStructure(True) Then
                 Module_Dashboard.DebugLog "GetOrCreateSheet", "WARNING - Failed to re-protect workbook structure after adding sheet '" & ws.Name & "'."
             Else
                 Module_Dashboard.DebugLog "GetOrCreateSheet", "Structure re-locked."
             End If
         End If
         Module_Dashboard.DebugLog "GetOrCreateSheet", "Finished creating sheet '" & ws.Name & "'."
    Else
         Module_Dashboard.DebugLog "GetOrCreateSheet", "Worksheet '" & sName & "' already exists."
         ' Optional: Clear contents if sheet already exists?
         ' Module_Dashboard.DebugLog "GetOrCreateSheet", "Clearing existing sheet contents..."
         ' ws.Cells.Clear ' Uncomment if desired behavior
    End If

    ' --- Ensure sheet is visible ---
    If Not ws Is Nothing Then
         If ws.Visible <> xlSheetVisible Then
              Module_Dashboard.DebugLog "GetOrCreateSheet", "Sheet '" & ws.Name & "' was not visible. Setting Visible=True."
              ws.Visible = xlSheetVisible
         End If
    Else
         Module_Dashboard.DebugLog "GetOrCreateSheet", "ERROR - ws object is Nothing before visibility check."
         GoTo GetOrCreateSheet_Cleanup ' Should not happen if error handling is correct
    End If


    ' --- Apply title banner (A1:N1) ---
    Module_Dashboard.DebugLog "GetOrCreateSheet", "Applying title banner formatting..."
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
        Module_Dashboard.DebugLog "GetOrCreateSheet", "Warning - could not apply title formatting to '" & ws.Name & "'. Err=" & Err.Number & ": " & Err.Description
        Err.Clear
    End If
    On Error GoTo GetOrCreateSheet_Error ' Restore handler
    Module_Dashboard.DebugLog "GetOrCreateSheet", "Title banner applied."

GetOrCreateSheet_Cleanup:
    Set GetOrCreateSheet = ws ' Return the sheet object (might be Nothing if error occurred)
    Module_Dashboard.DebugLog "GetOrCreateSheet", "EXIT. Returning sheet object (Is Nothing = " & (ws Is Nothing) & "). Time: " & Format(Timer - t1, "0.00") & "s"
    Exit Function ' Normal exit

GetOrCreateSheet_Error:
    Module_Dashboard.DebugLog "GetOrCreateSheet", "!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!"
    Module_Dashboard.DebugLog "GetOrCreateSheet", "ERROR Handler! Err=" & Err.Number & ": " & Err.Description & " near line " & Erl
    Module_Dashboard.DebugLog "GetOrCreateSheet", "!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!"
    ' Attempt to re-lock structure if we unlocked it and an error occurred
    If wbStructureWasLocked Then Module_Dashboard.ToggleWorkbookStructure (True)
    Set ws = Nothing ' Ensure Nothing is returned on error
    Resume GetOrCreateSheet_Cleanup ' Go to cleanup to return Nothing

End Function

'--------------------------------------------------
'  Post-copy formatting + nav buttons + protection for Views
'  REVISED: Includes Row Height Fix
'--------------------------------------------------
Private Sub ApplyViewFormatting(ws As Worksheet, viewTag As String)
    Dim lastRow As Long
    Dim lastColLetter As String: lastColLetter = Module_Dashboard.DB_COL_COMMENTS ' Should be "N"
    Dim phaseColLetter As String: phaseColLetter = Module_Dashboard.DB_COL_PHASE ' Should be "L"
    Dim workflowColLetter As String: workflowColLetter = Module_Dashboard.DB_COL_WORKFLOW_LOCATION ' Should be "J"
    Dim wsSrc As Worksheet
    Dim t1 As Double: t1 = Timer

    On Error GoTo ApplyViewFormatting_Error ' Use specific handler
    Module_Dashboard.DebugLog "ApplyViewFormatting", "ENTER. TargetSheet='" & ws.Name & "', ViewTag='" & viewTag & "'"

    If ws Is Nothing Then
        Module_Dashboard.DebugLog "ApplyViewFormatting", "ERROR: ws object Is Nothing on entry. Exiting."
        Exit Sub
    End If

    Application.ScreenUpdating = False ' Keep off during formatting

    ' --- Unprotect first ---
    Module_Dashboard.DebugLog "ApplyViewFormatting", "Unprotecting sheet..."
    On Error Resume Next ' Handle error if already unprotected or password wrong
    ws.Unprotect Password:=Module_Dashboard.PW_WORKBOOK ' Use password if defined
    If Err.Number <> 0 Then Module_Dashboard.DebugLog "ApplyViewFormatting", "Warning: Error unprotecting sheet. Err=" & Err.Number: Err.Clear
    On Error GoTo ApplyViewFormatting_Error ' Restore handler

    ' --- Force banner + control row heights to match main dashboard --- <<< ROW HEIGHT FIX ADDED >>>
    Module_Dashboard.DebugLog "ApplyViewFormatting", "Applying consistent row heights..."
    On Error Resume Next ' Ignore errors if setting height fails for any reason
    ws.rows(1).RowHeight = 32  ' Match Title Bar height from SetupDashboard
    ws.rows(2).RowHeight = 28  ' Match Control Panel height from SetupDashboard
    ' --- Optional: Set consistent data row height ---
    ' Need to calculate lastRow *before* setting data row height if moved here
    lastRow = ws.Cells(ws.rows.Count, "A").End(xlUp).Row ' Calculate lastRow earlier
    If lastRow >= 4 Then
        ws.rows("4:" & lastRow).RowHeight = 18 ' Example: Set data rows to 18 points
    End If
    If Err.Number <> 0 Then Module_Dashboard.DebugLog "ApplyViewFormatting", "Warning: Error setting row heights. Err=" & Err.Number: Err.Clear
    On Error GoTo ApplyViewFormatting_Error ' Restore main error handler
    Module_Dashboard.DebugLog "ApplyViewFormatting", "Applied consistent row heights for Rows 1, 2" & IIf(lastRow >= 4, " and Data Rows.", ".")
    ' --- End Row Height Fix ---

    ' --- Apply Column Widths (Copy from Source Dashboard) ---
    Module_Dashboard.DebugLog "ApplyViewFormatting", "Applying column widths..."
    On Error Resume Next
    Set wsSrc = ThisWorkbook.Worksheets(Module_Dashboard.DASHBOARD_SHEET_NAME) ' Uses original constant
    If Err.Number <> 0 Or wsSrc Is Nothing Then
        Module_Dashboard.DebugLog "ApplyViewFormatting", "Warning - Could not get source dashboard sheet '" & Module_Dashboard.DASHBOARD_SHEET_NAME & "' to copy widths. Err=" & Err.Number
        Err.Clear
        ' Fallback: AutoFit A:N if source widths unavailable
        Module_Dashboard.DebugLog "ApplyViewFormatting", "Applying AutoFit as fallback..."
        ws.Columns("A:" & lastColLetter).AutoFit
    Else
        wsSrc.Range("A1:" & lastColLetter & "1").Copy
        ws.Range("A1").PasteSpecial xlPasteColumnWidths
        Application.CutCopyMode = False
        Module_Dashboard.DebugLog "ApplyViewFormatting", "Copied column widths from '" & wsSrc.Name & "'."
    End If
    Set wsSrc = Nothing ' Release source sheet object
    On Error GoTo ApplyViewFormatting_Error ' Restore handler

    ' --- Apply Number Formatting ---
    Module_Dashboard.DebugLog "ApplyViewFormatting", "Applying number formats..."
    ' lastRow = ws.Cells(ws.rows.Count, "A").End(xlUp).Row ' Moved calculation earlier for Row Height fix
    Module_Dashboard.DebugLog "ApplyViewFormatting", "lastRow for formatting = " & lastRow
    ' Data starts at Row 4 on these views because we copy A1:N3 headers
    If lastRow >= 4 Then
        On Error Resume Next ' Handle errors applying formats
        ws.Range("D4:D" & lastRow).NumberFormat = "$#,##0.00"   ' Amount (D)
        ws.Range("E4:E" & lastRow).NumberFormat = "mm/dd/yyyy" ' Document Date (E)
        ws.Range("F4:F" & lastRow).NumberFormat = "mm/dd/yyyy" ' First Date Pulled (F)
        ws.Range("I4:I" & lastRow).NumberFormat = "0"           ' Pull Count (I)
        ws.Range(Module_Dashboard.DB_COL_LASTCONTACT & "4:" & Module_Dashboard.DB_COL_LASTCONTACT & lastRow).NumberFormat = "mm/dd/yyyy" ' Last Contact (M)
        If Err.Number <> 0 Then Module_Dashboard.DebugLog "ApplyViewFormatting", "Warning: Error applying number formats. Err=" & Err.Number: Err.Clear
        On Error GoTo ApplyViewFormatting_Error ' Restore handler
    Else
         Module_Dashboard.DebugLog "ApplyViewFormatting", "Skipping number formats (lastRow < 4)."
    End If

    ' --- Apply Conditional Formatting ---
    If lastRow >= 4 Then ' Data starts row 4
         Module_Dashboard.DebugLog "ApplyViewFormatting", "Applying conditional formatting..."
         Module_Dashboard.DebugLog "ApplyViewFormatting", "Calling ApplyColorFormatting..."
         Module_Dashboard.ApplyColorFormatting ws, 4 ' Start formatting data from row 4
         Module_Dashboard.DebugLog "ApplyViewFormatting", "Calling ApplyWorkflowLocationFormatting..."
         Module_Dashboard.ApplyWorkflowLocationFormatting ws, 4 ' Start formatting data from row 4
         Module_Dashboard.DebugLog "ApplyViewFormatting", "Finished applying conditional formatting."
    Else
         Module_Dashboard.DebugLog "ApplyViewFormatting", "Skipping conditional formatting (lastRow < 4)."
    End If

    ' --- Protection (Make read-only) ---
    Module_Dashboard.DebugLog "ApplyViewFormatting", "Locking all cells..."
    On Error Resume Next ' Handle error if sheet is protected
    ws.Cells.Locked = True ' Lock all cells on the view sheets
    If Err.Number <> 0 Then Module_Dashboard.DebugLog "ApplyViewFormatting", "Warning: Error locking cells. Err=" & Err.Number: Err.Clear
    On Error GoTo ApplyViewFormatting_Error ' Restore handler

    ' --- Apply Freeze Panes ---
    Module_Dashboard.DebugLog "ApplyViewFormatting", "Applying freeze panes..."
    On Error Resume Next ' Handle errors if sheet not active or already frozen/unfrozen
    ws.Activate ' Must activate to set freeze panes
    ActiveWindow.FreezePanes = False ' Unfreeze first
    ws.Range("A4").Select           ' Select cell below freeze rows (1-3)
    ActiveWindow.FreezePanes = True  ' Freeze Rows 1-3
    ws.Range("A1").Select ' Select A1 after freezing
    If Err.Number <> 0 Then Module_Dashboard.DebugLog "ApplyViewFormatting", "Warning: Error applying freeze panes. Err=" & Err.Number: Err.Clear
    On Error GoTo ApplyViewFormatting_Error ' Restore handler
    Module_Dashboard.DebugLog "ApplyViewFormatting", "Applied freeze panes."

     ' --- Final Protection ---
     Module_Dashboard.DebugLog "ApplyViewFormatting", "Applying final sheet protection..."
     On Error Resume Next ' Handle protection errors
    ws.Protect Password:=Module_Dashboard.PW_WORKBOOK, DrawingObjects:=True, Contents:=True, Scenarios:=True, _
        AllowFormattingCells:=False, AllowFormattingColumns:=False, AllowFormattingRows:=False, _
        AllowInsertingColumns:=False, AllowInsertingRows:=False, AllowInsertingHyperlinks:=False, _
        AllowDeletingColumns:=False, AllowDeletingRows:=False, AllowSorting:=False, AllowFiltering:=False, _
        AllowUsingPivotTables:=False
    If Err.Number <> 0 Then Module_Dashboard.DebugLog "ApplyViewFormatting", "Warning: Error applying sheet protection. Err=" & Err.Number: Err.Clear
    On Error GoTo ApplyViewFormatting_Error ' Restore handler
    Module_Dashboard.DebugLog "ApplyViewFormatting", "Applied sheet protection (Read-Only)."

    ' --- Add Navigation Buttons ---
    Module_Dashboard.DebugLog "ApplyViewFormatting", "Calling AddNavigationButtons..."
    AddNavigationButtons ws ' Call helper to add buttons
    Module_Dashboard.DebugLog "ApplyViewFormatting", "Returned from AddNavigationButtons."

ApplyViewFormatting_Cleanup:
    Module_Dashboard.DebugLog "ApplyViewFormatting", "Cleanup Label Reached."
    Application.ScreenUpdating = True ' Restore screen updating
    Set wsSrc = Nothing
    Module_Dashboard.DebugLog "ApplyViewFormatting", "EXIT (Cleanup Complete). Time: " & Format(Timer - t1, "0.00") & "s"
    Exit Sub ' Normal Exit

ApplyViewFormatting_Error:
    Module_Dashboard.DebugLog "ApplyViewFormatting", "!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!"
    Module_Dashboard.DebugLog "ApplyViewFormatting", "ERROR Handler! Err=" & Err.Number & ": " & Err.Description & " near line " & Erl
    Module_Dashboard.DebugLog "ApplyViewFormatting", "!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!"
    Resume ApplyViewFormatting_Cleanup ' Go to cleanup even on error

End Sub

'--------------------------------------------------
'  AddNavigationButtons – Row-2 layout (C D F G H) + Placeholders J K L N
'--------------------------------------------------
Public Sub AddNavigationButtons(ws As Worksheet)

    Const TOPROW       As Long = 2
    Const BTN_H        As Double = 24     ' Matches height set in ModernButton
    Const BTN_W_STD    As Double = 110    ' Adjust width for C/D as needed (match ColWidth setting)
    Const BTN_W_NAV    As Double = 75     ' Adjust width for F/G/H as needed (match ColWidth setting)

    Dim btnDefs As Variant, def As Variant, shp As Shape, target As Range, i As Long
    Dim t1 As Double: t1 = Timer


    ' --- Define each button: { Target Cell Address, Caption, Macro Name, Width } ---
    btnDefs = Array( _
      Array("C2", "Standard Refresh", "Button_RefreshDashboard_SaveAndRestoreEdits", 0), _
      Array("D2", "Preserve UserEdits", "Button_RefreshDashboard_PreserveUserEdits", 0), _
      Array("F2", "All Items", "modArchival.btnViewAll", 65), _
      Array("G2", "Active", "modArchival.btnViewActive", 65), _
      Array("H2", "Archive", "modArchival.btnViewArchive", 65) _
    )
    ' Element Definitions:
    ' Index 0: Target Cell Address (String) - e.g., "C2"
    ' Index 1: Button Caption (String) - e.g., "Standard Refresh"
    ' Index 2: Macro Name to Run (String) - e.g., "Button_RefreshDashboard_SaveAndRestoreEdits"
    ' Index 3: Button Width (Double) - Use 0 for AutoFit-to-Cell, or specific width like 65
    
    On Error GoTo AddNav_Error
    Module_Dashboard.DebugLog "AddNavigationButtons", "ENTER for sheet: '" & ws.Name & "' - Layout C2/D2/F2/G2/H2"

    If ws Is Nothing Then Exit Sub

    '----- preparation: Unprotect, Clear Row 2 Shapes/Content/Merges -----
    On Error Resume Next
    ws.Unprotect Password:=Module_Dashboard.PW_WORKBOOK
    Dim deleteCount As Long: deleteCount = 0
    For Each shp In ws.Shapes ' Delete ANY shape anchored in Row 2
        If shp.TopLeftCell.Row = TOPROW Then
             If shp.Name Like "btn_*" Or shp.Name Like "nav*" Or shp.Type = msoFormControl Then
                shp.Delete
                deleteCount = deleteCount + 1
             End If
        End If
    Next shp
    ws.Range("B" & TOPROW & ":N" & TOPROW).ClearContents ' Clear B2:N2
    ws.Range("B" & TOPROW & ":N" & TOPROW).UnMerge       ' Unmerge all relevant cells just in case
    If Err.Number <> 0 Then Module_Dashboard.DebugLog "AddNavigationButtons", "Warning during clear/unmerge: Err=" & Err.Number: Err.Clear
    On Error GoTo AddNav_Error
    Module_Dashboard.DebugLog "AddNavigationButtons", "Cleared Row " & TOPROW & ". Deleted " & deleteCount & " shapes."

    '----- create the five buttons ------------------------------------
    Module_Dashboard.DebugLog "AddNavigationButtons", "Creating buttons..."
    For i = LBound(btnDefs) To UBound(btnDefs)
        def = btnDefs(i)
        Set target = Nothing
        On Error Resume Next
        Set target = ws.Range(def(0)) ' Get target Cell (e.g., C2)
        On Error GoTo AddNav_Error

        If target Is Nothing Then
            Module_Dashboard.DebugLog "AddNavigationButtons", "ERROR: Invalid target cell '" & def(0) & "'. Skipping."
        Else
            Set shp = Nothing
            ' Call ModernButton, passing width from def(3)
            Set shp = Module_Dashboard.ModernButton(ws, target, def(1), def(2), def(3))

            If Not shp Is Nothing Then
            
                ' --- START: ADD Height Adjustment ---
                ' Make Nav buttons shorter for sleeker look
                If def(0) = "F2" Or def(0) = "G2" Or def(0) = "H2" Then
                    shp.Height = 18                ' Use shorter height (e.g., 18)
                    Module_Dashboard.DebugLog "AddNavigationButtons", "Adjusted height to 18 for: " & def(1)
                ' Optional: Ensure other buttons retain standard height if needed, though ModernButton should handle default
                ' Else
                '    shp.Height = BTN_H ' BTN_H should be defined above, e.g., 24
                End If
                ' --- END: ADD Height Adjustment ---
                
            ' --- START: ADD THIS IF BLOCK ---
            ' Auto-size width if definition used 0
            If def(3) = 0 Then                ' Check if width is the AutoFit flag
                On Error Resume Next ' Handle potential error reading target width
                shp.Width = target.Width - 6  ' Set width based on cell, minus padding
                If Err.Number <> 0 Then       ' Fallback if error reading width
                    shp.Width = 110 ' Default wide width if calculation fails
                    Err.Clear
                End If
                On Error GoTo AddNav_Error ' Restore main handler
            End If
            ' --- END: ADD THIS IF BLOCK ---

                On Error Resume Next ' Handle errors modifying shape
                shp.Name = "btn_" & Replace(def(1), " ", "_") ' Set final name
                 ' Center the button (ModernButton might already approximate this)
                 ' Recalculate here for precision based on actual cell/button dimensions
                shp.Left = target.Left + (target.Width - shp.Width) / 2
                shp.Top = target.Top + (target.Height - shp.Height) / 2
                If Err.Number <> 0 Then Module_Dashboard.DebugLog "AddNavigationButtons", "Warning: Error centering/naming '" & shp.Name & "'. Err=" & Err.Number: Err.Clear
                On Error GoTo AddNav_Error
            Else
                 Module_Dashboard.DebugLog "AddNavigationButtons", "ERROR: ModernButton failed for '" & def(1) & "'."
            End If
        End If
    Next i
    Module_Dashboard.DebugLog "AddNavigationButtons", "Finished creating buttons."

    '----- placeholder counts & timestamp -----------------------------
    Module_Dashboard.DebugLog "AddNavigationButtons", "Setting placeholder counts/timestamp..."
    On Error Resume Next
    With ws.Range("J" & TOPROW & ":L" & TOPROW) ' J2:L2
        .Value = Array("All: TBC", "Active: TBC", "Archive: TBC") ' Placeholders
        .Font.Size = 9
        .Font.Italic = True
        .VerticalAlignment = xlCenter
        .HorizontalAlignment = xlLeft ' Align left
    End With
    With ws.Range("N" & TOPROW) ' N2
        .Value = "Refreshed: " & Format$(Now(), "mm/dd hh:nn AM/PM")
        .Font.Size = 9
        .Font.Italic = True
        .HorizontalAlignment = xlLeft ' Align left
        .VerticalAlignment = xlCenter
    End With
    If Err.Number <> 0 Then Module_Dashboard.DebugLog "AddNavigationButtons", "Warning: Error setting counts/timestamp text. Err=" & Err.Number: Err.Clear
    On Error GoTo AddNav_Error

    '----- re-protect --------------------------------------------------
    Module_Dashboard.DebugLog "AddNavigationButtons", "Re-protecting sheet..."
    On Error Resume Next
    ws.Protect Password:=Module_Dashboard.PW_WORKBOOK, DrawingObjects:=True, Contents:=True, Scenarios:=True
    If Err.Number <> 0 Then Module_Dashboard.DebugLog "AddNavigationButtons", "Warning: Error re-protecting. Err=" & Err.Number: Err.Clear
    On Error GoTo AddNav_Error
    Module_Dashboard.DebugLog "AddNavigationButtons", "Sheet re-protected."

    Module_Dashboard.DebugLog "AddNavigationButtons", "EXIT (Normal - New Layout). Time: " & Format(Timer - t1, "0.00") & "s"
    Exit Sub

AddNav_Error:
    Module_Dashboard.DebugLog "AddNavigationButtons", "!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!"
    Module_Dashboard.DebugLog "AddNavigationButtons", "ERROR Handler! Sheet='" & ws.Name & "'. Err=" & Err.Number & ": " & Err.Description & " near line " & Erl
    Module_Dashboard.DebugLog "AddNavigationButtons", "!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!"
    Set shp = Nothing: Set target = Nothing
End Sub


