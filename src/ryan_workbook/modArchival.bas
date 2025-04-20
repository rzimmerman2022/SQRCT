'=====================================================================
' Module :  modArchival   – Archive / Active / All dashboard views
' Purpose:  Handles the creation, filtering, formatting, and navigation
'           for the filtered 'SQRCT Active' and 'SQRCT Archive' views,
'           based on data from the main 'SQRCT Dashboard'. Also contains
'           logic for storing record counts via properties.
' Workbook: SQRCT
' REVISED:  CORRECTED MODULE STRUCTURE (Declarations first). Added count properties,
'           updated subs to calculate/store counts and remove TBC placeholders.
'           Incorporated fixes from previous versions (Row Height, Blank Phase=Active,
'           Naming Convention, ReDim Preserve fix, detailed logging).
'           Refined section structure for better organization.
'           FIXED: Removed invalid 'Me' keyword usage when calling Property Let within the standard module.
' REVISED: 04/20/2025 - Refactored ApplyViewFormatting to call new
'                      modFormatting.ExactlyCloneDashboardFormatting function
'                      to ensure consistent Row 1/2 appearance. Removed
'                      premature sheet protection from AddNavigationButtons.
'=====================================================================
Option Explicit

'--------------------------------------------------
' SECTION 1: MODULE-LEVEL DECLARATIONS
'--------------------------------------------------
' ***** All Const declarations MUST come before any Subs, Functions, Properties, or Private variable declarations *****

' --- Phase Categorization ---
' Defines which phases are considered "Active" for filtering purposes.
' Pipe delimiters ensure exact, case-insensitive matching.
Private Const ACTIVE_PHASES As String = "|FIRST F/U|AF|RZ|KMH|RI|OTHER (ACTIVE)|" ' Added leading/trailing pipes

' --- Sheet Names & UI Text ---
' Defines standard names and titles for the sheets managed by this module.
' Public constants allow other modules (like modUtilities) to refer to these sheets reliably.
Public Const SH_ACTIVE As String = "SQRCT Active"      ' Public Name for the Active sheet
Public Const SH_ARCHIVE As String = "SQRCT Archive"    ' Public Name for the Archive sheet
Private Const TITLE_ACTIVE As String = _
        "STRATEGIC QUOTE RECOVERY & CONVERSION TRACKER – ACTIVE VIEW" ' Sheet title banner text
Private Const TITLE_ARCHIVE As String = _
        "STRATEGIC QUOTE RECOVERY & CONVERSION TRACKER – ARCHIVE VIEW" ' Sheet title banner text
' (Add any other module-level Const declarations needed for this module right here)


' --- Module-Level Count Storage (Backing Fields) ---
' Private variables to hold the counts calculated during the refresh process.
' Accessed externally via Property Procedures in Section 2.
Private m_totalRecords As Long    ' Stores total records from main dashboard (calculated in RefreshAllViews)
Private m_activeRecords As Long   ' Stores count of records filtered into Active view (calculated in CopyFilteredRows)
Private m_archiveRecords As Long  ' Stores count of records filtered into Archive view (calculated in CopyFilteredRows)
' (Add any other module-level Dim or Private variable declarations right here)


'--------------------------------------------------
' SECTION 2: COUNT PROPERTY PROCEDURES
'--------------------------------------------------
' Public properties provide controlled, read/write access to the count values
' stored in the m_... backing fields declared in Section 1.
' Allows other modules (e.g., modUtilities or Module_Dashboard) to get the counts
' after they have been calculated by this module's refresh process.

Public Property Get TotalRecords() As Long
    ' Returns the current value of the total record count backing field.
    TotalRecords = m_totalRecords
End Property
Public Property Let TotalRecords(value As Long)
    ' Sets the value of the total record count backing field.
    ' Called By: RefreshAllViews
    m_totalRecords = value
End Property

Public Property Get ActiveRecords() As Long
    ' Returns the current value of the active record count backing field.
    ActiveRecords = m_activeRecords
End Property
Public Property Let ActiveRecords(value As Long)
    ' Sets the value of the active record count backing field.
    ' Called By: CopyFilteredRows (when keepArchived=False)
    m_activeRecords = value
End Property

Public Property Get ArchiveRecords() As Long
    ' Returns the current value of the archive record count backing field.
    ArchiveRecords = m_archiveRecords
End Property
Public Property Let ArchiveRecords(value As Long)
    ' Sets the value of the archive record count backing field.
    ' Called By: CopyFilteredRows (when keepArchived=True)
    m_archiveRecords = value
End Property


'--------------------------------------------------
' SECTION 3: PUBLIC INTERFACE / BUTTON HANDLERS
'--------------------------------------------------
' Subroutines called directly by buttons assigned on the worksheets.
' These typically orchestrate calling the refresh logic and activating the relevant sheet.

Public Sub btnViewActive()
    ' Action for the "Active" view button. Refreshes the view THEN activates it.
    Module_Dashboard.DebugLog "btnViewActive", "ENTER"
    Application.StatusBar = "Refreshing Active View..." ' User feedback
    RefreshAndActivate SH_ACTIVE ' Call helper to refresh and activate the Active sheet by name
    Application.StatusBar = False ' Clear status bar
    Module_Dashboard.DebugLog "btnViewActive", "EXIT"
End Sub

Public Sub btnViewArchive()
    ' Action for the "Archive" view button. Refreshes the view THEN activates it.
    Module_Dashboard.DebugLog "btnViewArchive", "ENTER"
    Application.StatusBar = "Refreshing Archive View..." ' User feedback
    RefreshAndActivate SH_ARCHIVE ' Call helper to refresh and activate the Archive sheet by name
    Application.StatusBar = False ' Clear status bar
    Module_Dashboard.DebugLog "btnViewArchive", "EXIT"
End Sub

Public Sub btnViewAll()
    ' Action for the "All Items" view button. Simply activates the main dashboard sheet.
    Module_Dashboard.DebugLog "btnViewAll", "ENTER"
    Application.StatusBar = "Activating Main Dashboard..." ' User feedback
    On Error Resume Next ' Handle error if sheet doesn't exist or is hidden
    ' Uses Public Const from Module_Dashboard
    ThisWorkbook.Worksheets(Module_Dashboard.DASHBOARD_SHEET_NAME).Activate ' Activate the main dashboard sheet
    If Err.Number <> 0 Then ' Check if activation failed
         Module_Dashboard.DebugLog "btnViewAll", "ERROR activating sheet '" & Module_Dashboard.DASHBOARD_SHEET_NAME & "'. Err=" & Err.Number & ": " & Err.Description ' Log error
         MsgBox "Could not activate sheet: " & Module_Dashboard.DASHBOARD_SHEET_NAME, vbExclamation ' Inform user
         Err.Clear ' Clear the error
    Else
         Module_Dashboard.DebugLog "btnViewAll", "Activated sheet '" & Module_Dashboard.DASHBOARD_SHEET_NAME & "'" ' Log success
    End If
    On Error GoTo 0 ' Restore default error handling
    Application.StatusBar = False ' Clear status bar
    Module_Dashboard.DebugLog "btnViewAll", "EXIT"
End Sub


'--------------------------------------------------
' SECTION 4: MAIN VIEW REFRESH ORCHESTRATION
'--------------------------------------------------
' Contains the primary public entry point called by other modules (e.g., Module_Dashboard)
' to trigger the refresh of the Active and Archive views.

Public Sub RefreshAllViews() ' No argument needed
    ' Purpose: Orchestrates the entire refresh process for both Active and Archive views.
    '          Calculates total records, calls individual view refresh routines, and handles errors.
    ' Called By: Module_Dashboard.RefreshDashboard (typically)
    Dim wsDash As Worksheet
    Dim originalStatusBar As Variant
    Dim t1 As Double: t1 = Timer

    On Error GoTo ArchivalErrorHandler ' Use a specific handler for this sub

    Module_Dashboard.DebugLog "RefreshAllViews", "ENTER (Using Property Let/Get Plan)"
    InitializeCounts ' <<< ACTION: Reset m_ count properties at the very start

    ' --- Store & Set ONLY StatusBar ---
    originalStatusBar = Application.StatusBar ' Store current status bar text
    Application.StatusBar = "Refreshing Active/Archive views..."
    Module_Dashboard.DebugLog "RefreshAllViews", "StatusBar set."

    ' --- Get Main Dashboard Reference ---
    Module_Dashboard.DebugLog "RefreshAllViews", "Getting main dashboard sheet reference ('" & Module_Dashboard.DASHBOARD_SHEET_NAME & "')..."
    Set wsDash = Nothing ' Initialize
    On Error Resume Next ' Temporarily ignore error getting sheet reference
    Set wsDash = ThisWorkbook.Worksheets(Module_Dashboard.DASHBOARD_SHEET_NAME) ' Get object for main dashboard
    On Error GoTo ArchivalErrorHandler ' Restore main handler

    If wsDash Is Nothing Then
        Module_Dashboard.DebugLog "RefreshAllViews", "ERROR: Main dashboard sheet '" & Module_Dashboard.DASHBOARD_SHEET_NAME & "' not found. Cannot refresh views."
        GoTo ArchivalCleanup ' Go to cleanup if dashboard sheet missing
    End If
    Module_Dashboard.DebugLog "RefreshAllViews", "Got main dashboard sheet: '" & wsDash.Name & "'"

    ' <<< ACTION: CALCULATE AND SET TOTAL COUNT PROPERTY >>>
    Dim lastRow As Long
    lastRow = 0
    On Error Resume Next ' In case sheet is empty or Col A fails finding last row
    lastRow = wsDash.Cells(wsDash.rows.Count, "A").End(xlUp).Row ' Find last used row based on Col A data
    On Error GoTo ArchivalErrorHandler ' Restore handler
    If lastRow >= 4 Then ' Check if there are any data rows (Data starts row 4 on dashboard)
        TotalRecords = lastRow - 3 ' *** FIXED: Removed Me. *** SET: Store count via Property Let (Subtract 3 header rows)
    Else
        TotalRecords = 0 ' *** FIXED: Removed Me. *** SET: Store zero if no data rows
    End If
    Module_Dashboard.DebugLog "RefreshAllViews", "Calculated & Set TotalRecords Property: " & TotalRecords ' Use Property Get here for logging

    ' --- Refresh Views (These call CopyFilteredRows which sets Active/Archive properties) ---
    Module_Dashboard.DebugLog "RefreshAllViews", "Calling RefreshActiveView..."
    RefreshActiveView wsDash ' Create/Refresh Active view sheet
    Module_Dashboard.DebugLog "RefreshAllViews", "Returned from RefreshActiveView."

    Module_Dashboard.DebugLog "RefreshAllViews", "Calling RefreshArchiveView..."
    RefreshArchiveView wsDash ' Create/Refresh Archive view sheet
    Module_Dashboard.DebugLog "RefreshAllViews", "Returned from RefreshArchiveView."

    ' <<< NOTE: Count update UI (writing to J2:L2) is now called from RefreshDashboard (or modUtilities) AFTER this sub completes >>>
    ' <<< NOTE: AddNavigationButtons for wsDash is also called from RefreshDashboard AFTER this sub completes >>>

ArchivalCleanup:   ' Cleanup label for both normal exit and error
    Module_Dashboard.DebugLog "RefreshAllViews", "Cleanup Label Reached..."
    On Error Resume Next ' Ignore errors during cleanup itself
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


'--------------------------------------------------
' SECTION 5: INITIALIZATION & STATE MANAGEMENT
'--------------------------------------------------
' Contains routines related to setting up or resetting module state.

Private Sub InitializeCounts()
    ' Purpose: Resets all count backing fields to zero before a refresh.
    ' Called By: RefreshAllViews
    Module_Dashboard.DebugLog "InitializeCounts", "Resetting m_ count properties to 0." ' Log action
    m_totalRecords = 0   ' Reset total count backing field
    m_activeRecords = 0  ' Reset active count backing field
    m_archiveRecords = 0 ' Reset archive count backing field
End Sub


'--------------------------------------------------
' SECTION 6: VIEW-SPECIFIC REFRESH LOGIC
'--------------------------------------------------
' Contains routines that manage the refresh process for individual views (Active/Archive)
' and handle the activation of sheets.

Private Sub RefreshAndActivate(viewName As String) ' viewName will be SH_ACTIVE or SH_ARCHIVE
    ' Purpose: Refreshes a specific view (Active or Archive) and then activates its sheet.
    '          Handles getting the source data sheet and calling the appropriate refresh helper.
    ' Called By: btnViewActive, btnViewArchive
    Dim wsDash As Worksheet
    Dim originalStatusBar As Variant
    Dim t1 As Double: t1 = Timer

    On Error GoTo RefreshActivateErrorHandler ' Specific handler for this sub

    Module_Dashboard.DebugLog "RefreshAndActivate", "ENTER for viewName='" & viewName & "'"

    ' --- Store & Set ONLY StatusBar ---
    originalStatusBar = Application.StatusBar
    Application.StatusBar = "Refreshing " & viewName & "..." ' Update status
    Module_Dashboard.DebugLog "RefreshAndActivate", "StatusBar set."

    ' --- Get Main Dashboard Reference ---
    Module_Dashboard.DebugLog "RefreshAndActivate", "Getting main dashboard sheet reference ('" & Module_Dashboard.DASHBOARD_SHEET_NAME & "')..."
    Set wsDash = Nothing
    On Error Resume Next
    Set wsDash = ThisWorkbook.Worksheets(Module_Dashboard.DASHBOARD_SHEET_NAME) ' Uses public constant from Module_Dashboard
    On Error GoTo RefreshActivateErrorHandler ' Restore handler

    If wsDash Is Nothing Then
        Module_Dashboard.DebugLog "RefreshAndActivate", "ERROR in RefreshAndActivate: Main dashboard sheet not found."
        GoTo RefreshActivateCleanup ' Go to cleanup
    End If
    Module_Dashboard.DebugLog "RefreshAndActivate", "Got main dashboard sheet: '" & wsDash.Name & "'"

    ' --- Refresh the specific view ---
    Module_Dashboard.DebugLog "RefreshAndActivate", "Selecting view to refresh based on viewName='" & viewName & "'"
    Select Case viewName
        Case SH_ACTIVE ' Use constant defined in Section 1
            Module_Dashboard.DebugLog "RefreshAndActivate", "Calling RefreshActiveView..."
            RefreshActiveView wsDash
            Module_Dashboard.DebugLog "RefreshAndActivate", "Returned from RefreshActiveView."
        Case SH_ARCHIVE ' Use constant defined in Section 1
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
    ThisWorkbook.Worksheets(viewName).Activate ' Uses viewName directly (e.g. "SQRCT Active")
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

Private Sub RefreshActiveView(wsSrc As Worksheet) ' Takes source worksheet as parameter
    ' Purpose: Manages the refresh process specifically for the 'SQRCT Active' sheet.
    '          Gets/Creates the target sheet, calls CopyFilteredRows, and applies formatting.
    ' Called By: RefreshAllViews, RefreshAndActivate
    Dim wsTgt As Worksheet
    Dim t1 As Double: t1 = Timer

    On Error GoTo RefreshActive_Error ' Add error handling for this sub
    Module_Dashboard.DebugLog "RefreshActiveView", "ENTER. Source sheet='" & wsSrc.Name & "'"

    Module_Dashboard.DebugLog "RefreshActiveView", "Calling GetOrCreateSheet for '" & SH_ACTIVE & "'..." ' Use constant
    Set wsTgt = GetOrCreateSheet(SH_ACTIVE, TITLE_ACTIVE, RGB(0, 110, 0)) ' Dark Green Banner

    If wsTgt Is Nothing Then
        Module_Dashboard.DebugLog "RefreshActiveView", "GetOrCreateSheet failed. Aborting."
        GoTo RefreshActive_Cleanup ' Go to cleanup
    End If
    Module_Dashboard.DebugLog "RefreshActiveView", "Got target sheet: '" & wsTgt.Name & "'"

    Module_Dashboard.DebugLog "RefreshActiveView", "Calling CopyFilteredRows (keepArchived=False)..."
    CopyFilteredRows wsSrc, wsTgt, keepArchived:=False ' Filter for Active phases
    Module_Dashboard.DebugLog "RefreshActiveView", "Returned from CopyFilteredRows."

    ' --- Add Defensive Check before Formatting ---
    Module_Dashboard.DebugLog "RefreshActiveView", "Performing defensive checks before formatting..."
    If wsTgt Is Nothing Then ' Double check object variable just in case
         Module_Dashboard.DebugLog "RefreshActiveView", "Defensive Check FAIL: wsTgt object became Nothing unexpectedly."
    ElseIf Not SheetExists(wsTgt.Name) Then ' Check if sheet still exists by name
         Module_Dashboard.DebugLog "RefreshActiveView", "Defensive Check FAIL: Target sheet '" & SH_ACTIVE & "' no longer exists before formatting." ' Use constant
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
    ' --- Re-enable Events if they were turned off (though this sub doesn't disable them) ---
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

Private Sub RefreshArchiveView(wsSrc As Worksheet) ' Takes source worksheet as parameter
    ' Purpose: Manages the refresh process specifically for the 'SQRCT Archive' sheet.
    '          Gets/Creates the target sheet, calls CopyFilteredRows, and applies formatting.
    ' Called By: RefreshAllViews, RefreshAndActivate
    Dim wsTgt As Worksheet
    Dim t1 As Double: t1 = Timer

    On Error GoTo RefreshArchive_Error ' Add error handling for this sub
    Module_Dashboard.DebugLog "RefreshArchiveView", "ENTER. Source sheet='" & wsSrc.Name & "'"

    Module_Dashboard.DebugLog "RefreshArchiveView", "Calling GetOrCreateSheet for '" & SH_ARCHIVE & "'..." ' Use constant
    Set wsTgt = GetOrCreateSheet(SH_ARCHIVE, TITLE_ARCHIVE, RGB(150, 40, 40)) ' Dark Red Banner

    If wsTgt Is Nothing Then
        Module_Dashboard.DebugLog "RefreshArchiveView", "GetOrCreateSheet failed. Aborting."
        GoTo RefreshArchive_Cleanup ' Go to cleanup
    End If
    Module_Dashboard.DebugLog "RefreshArchiveView", "Got target sheet: '" & wsTgt.Name & "'"

    Module_Dashboard.DebugLog "RefreshArchiveView", "Calling CopyFilteredRows (keepArchived=True)..." ' Filter for Archived phases
    CopyFilteredRows wsSrc, wsTgt, keepArchived:=True
    Module_Dashboard.DebugLog "RefreshArchiveView", "Returned from CopyFilteredRows."

    ' --- Add Defensive Check before Formatting (Similar to RefreshActiveView) ---
    Module_Dashboard.DebugLog "RefreshArchiveView", "Performing defensive checks before formatting..."
    If wsTgt Is Nothing Then
        Module_Dashboard.DebugLog "RefreshArchiveView", "Defensive Check FAIL: wsTgt object became Nothing unexpectedly."
    ElseIf Not SheetExists(wsTgt.Name) Then
        Module_Dashboard.DebugLog "RefreshArchiveView", "Defensive Check FAIL: Target sheet '" & SH_ARCHIVE & "' no longer exists before formatting."
    Else
        Module_Dashboard.DebugLog "RefreshArchiveView", "Defensive Check PASS: Target sheet valid, calling ApplyViewFormatting..."
        ApplyViewFormatting wsTgt, "Archive" ' Apply formatting and buttons
        Module_Dashboard.DebugLog "RefreshArchiveView", "Returned from ApplyViewFormatting."
    End If
    ' --- End Defensive Check ---

    Module_Dashboard.DebugLog "RefreshArchiveView", "Process completed normally (before cleanup)."

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
' SECTION 7: CORE FILTERING & DATA COPYING
'--------------------------------------------------
' Contains the main data processing logic: reading data into arrays,
' filtering based on phase, and writing results to the target sheet.
' Also includes the phase checking helper functions.

Private Sub CopyFilteredRows(wsSrc As Worksheet, wsTgt As Worksheet, _
                             keepArchived As Boolean)
    ' Purpose: Reads data from source sheet (main dashboard) into an array,
    '          filters rows based on the 'keepArchived' flag using IsPhaseArchived/IsPhaseActive,
    '          copies matching rows to the target sheet (Active or Archive view),
    '          and updates the corresponding module-level count property (ActiveRecords or ArchiveRecords).
    ' Called By: RefreshActiveView, RefreshArchiveView
    Dim lastSrcRow As Long, destRow As Long, r As Long
    Dim phaseValue As String
    Dim arrData As Variant       ' Array to hold source data A:N
    Dim i As Long, j As Long     ' Loop counters
    Dim outputData() As Variant  ' Array to hold rows TO BE copied (potentially oversized)
    Dim finalOutputData() As Variant ' Array sized exactly for output
    Dim outputRowCount As Long
    Dim srcRange As Range
    Dim phaseColIndex As Long
    Dim numCols As Long
    Dim t1 As Double: t1 = Timer

    Module_Dashboard.DebugLog "CopyFilteredRows", "ENTER. Source='" & wsSrc.Name & "', Target='" & wsTgt.Name & "', KeepArchived=" & keepArchived

    ' --- Define Source Data Range (A to N) ---
    Dim lastColLetter As String: lastColLetter = Module_Dashboard.DB_COL_COMMENTS ' Should be "N" (from Module_Dashboard)
    Dim firstDataRow As Long: firstDataRow = 4 ' Data starts at row 4 on main dash
    Dim expectedCols As Long: expectedCols = 14 ' Expected columns A-N

    On Error GoTo CopyFilteredRows_Error ' Add local error handler

    ' --- Ensure Target Sheet is Writable ---
    Module_Dashboard.DebugLog "CopyFilteredRows", "Ensuring target sheet '" & wsTgt.Name & "' is writable..."
    On Error Resume Next ' Handle error if already unprotected or password wrong
    If wsTgt.ProtectContents Then
        Module_Dashboard.DebugLog "CopyFilteredRows", "Target sheet is protected. Attempting unprotect..."
        wsTgt.Unprotect Password:=Module_Dashboard.PW_WORKBOOK ' Use constant from Module_Dashboard
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

        ' <<< ACTION: Ensure NO count is written here >>>
        ' The block that previously wrote "Count: 0" to J2 has been removed.
        ' Final counts will be written later by the calling module after reading properties.
        ' Set the relevant count property to 0 for this view
        If keepArchived Then
            ArchiveRecords = 0 ' *** FIXED: Removed Me. ***
            Module_Dashboard.DebugLog "CopyFilteredRows", "Set ArchiveRecords Property: 0 (No data)"
        Else
            ActiveRecords = 0 ' *** FIXED: Removed Me. ***
            Module_Dashboard.DebugLog "CopyFilteredRows", "Set ActiveRecords Property: 0 (No data)"
        End If

        Module_Dashboard.DebugLog "CopyFilteredRows", "EXIT (No data case). Time: " & Format(Timer - t1, "0.00") & "s"
        Exit Sub
    End If

    ' --- Read source data into array for faster processing ---
    Module_Dashboard.DebugLog "CopyFilteredRows", "Reading source range A" & firstDataRow & ":" & lastColLetter & lastSrcRow & " into array..."
    On Error Resume Next
    Set srcRange = wsSrc.Range("A" & firstDataRow & ":" & lastColLetter & lastSrcRow) ' A4:N<lastRow>
    arrData = srcRange.Value2 ' Use Value2 for potentially faster read and cleaner data (no currency/date formatting)
    If Err.Number <> 0 Then
         Module_Dashboard.DebugLog "CopyFilteredRows", "ERROR reading range into arrData. Err=" & Err.Number & ": " & Err.Description
         Err.Clear ' Clear error before checking IsArray
    End If
    On Error GoTo CopyFilteredRows_Error ' Restore Handler

    ' Check if reading failed or didn't produce an array
    If Not IsArray(arrData) Then
        Module_Dashboard.DebugLog "CopyFilteredRows", "ERROR: arrData is not an array after reading range. Exiting."
        ' Set the relevant count property to 0 as we can't process
        If keepArchived Then ArchiveRecords = 0 Else ActiveRecords = 0 ' *** FIXED: Removed Me. ***
        Exit Sub ' Or handle error appropriately
    End If
    Module_Dashboard.DebugLog "CopyFilteredRows", "Read successful. UBound(arrData, 1)=" & UBound(arrData, 1)

    ' >>> Array Dimension Handling (Check actual columns read) <<<
    Module_Dashboard.DebugLog "CopyFilteredRows", "Checking actual array dimensions read..."
    Dim actualCols As Long
    On Error Resume Next ' Handle error getting UBound if arrData is invalid (e.g., only 1 cell read)
    actualCols = UBound(arrData, 2)
    If Err.Number <> 0 Then
        ' Could be a 1D array (single row read) or actual error
        Err.Clear ' Clear error before checking 1D possibility
        On Error Resume Next
        actualCols = UBound(arrData, 1) ' Check if it's 1D by getting UBound of 1st dim
        If Err.Number = 0 And LBound(arrData, 1) = 1 Then ' Check if 1D array with LBound 1
             numCols = actualCols ' If 1D, numCols is the number of items
             Module_Dashboard.DebugLog "CopyFilteredRows", "Detected 1D array (single row). numCols set to " & numCols
        Else ' Genuine error or unexpected array structure
             Module_Dashboard.DebugLog "CopyFilteredRows", "ERROR getting UBound(arrData, 2) or invalid array structure. Setting numCols=0. Err=" & Err.Number & ": " & Err.Description
             numCols = 0 ' Handle error getting UBound
             Err.Clear
        End If
    Else
        numCols = actualCols ' Use the actual column count read from 2nd dimension
    End If
    On Error GoTo CopyFilteredRows_Error ' Restore Handler
    Module_Dashboard.DebugLog "CopyFilteredRows", "Actual columns read = " & actualCols & ". numCols set to " & numCols

    ' --- Handle Single Row Case (Value2 read might give 1D array) ---
    ' This block converts a 1D array (from reading a single row) into a 2D (1 row x N cols) array
    If lastSrcRow = firstDataRow Then ' Check if only one data row was expected
         Dim is2D As Boolean
         On Error Resume Next
         Dim dummy As Long: dummy = UBound(arrData, 2) ' Check if 2nd dimension exists
         is2D = (Err.Number = 0)
         On Error GoTo CopyFilteredRows_Error ' Restore Handler

         If Not is2D Then ' It's 1D
             Module_Dashboard.DebugLog "CopyFilteredRows", "Handling single row case (1D array found)."
             Dim tempArr() As Variant
             Dim itemsIn1D As Long
             On Error Resume Next ' Handle error getting bounds of 1D array
             itemsIn1D = UBound(arrData, 1) - LBound(arrData, 1) + 1
             If Err.Number <> 0 Then
                Module_Dashboard.DebugLog "CopyFilteredRows", "ERROR getting bounds of 1D array. Exiting."
                If keepArchived Then ArchiveRecords = 0 Else ActiveRecords = 0 ' *** FIXED: Removed Me. ***
                Exit Sub
             End If
             On Error GoTo CopyFilteredRows_Error ' Restore Handler

             numCols = itemsIn1D ' Number of columns is number of items in 1D array
             Module_Dashboard.DebugLog "CopyFilteredRows", "1D array has " & itemsIn1D & " items. Setting numCols=" & numCols
             ReDim tempArr(1 To 1, 1 To numCols) ' Size 2D array based on actual items read

             ' Copy elements from 1D to 2D array
             For j = LBound(arrData, 1) To UBound(arrData, 1)
                 tempArr(1, j) = arrData(j)
             Next j
             Erase arrData ' Clear original 1D array
             arrData = tempArr ' Replace original 1D array with 2D version
             Module_Dashboard.DebugLog "CopyFilteredRows", "Converted 1D to 2D array (1x" & numCols & ")."
         End If
    End If

    ' --- Prepare initial (potentially oversized) output array ---
    If numCols = 0 Then ' Safety check if numCols failed to be set
        Module_Dashboard.DebugLog "CopyFilteredRows", "ERROR: numCols is 0 before preparing output array. Exiting."
        If keepArchived Then ArchiveRecords = 0 Else ActiveRecords = 0 ' *** FIXED: Removed Me. ***
        Exit Sub
    End If
    Module_Dashboard.DebugLog "CopyFilteredRows", "Preparing initial output array: 1 To " & UBound(arrData, 1) & ", 1 To " & numCols
    ReDim outputData(1 To UBound(arrData, 1), 1 To numCols) ' Max possible size based on source rows
    outputRowCount = 0

    ' --- Get the 1-based index for the Phase column (L) within the array ---
    On Error Resume Next ' In case constant refers to invalid column
    phaseColIndex = wsSrc.Columns(Module_Dashboard.DB_COL_PHASE).Column ' Get numeric index for Phase (e.g., 12 for L)
    If Err.Number <> 0 Or phaseColIndex = 0 Or phaseColIndex > numCols Then
        Module_Dashboard.DebugLog "CopyFilteredRows", "ERROR - Invalid Phase Column '" & Module_Dashboard.DB_COL_PHASE & "' index (" & phaseColIndex & "). numCols=" & numCols
        Err.Clear ' Clear error before exiting
        Erase arrData ' Clean up input array
        If keepArchived Then ArchiveRecords = 0 Else ActiveRecords = 0 ' *** FIXED: Removed Me. ***
        Exit Sub
    End If
    On Error GoTo CopyFilteredRows_Error ' Restore Handler
    Module_Dashboard.DebugLog "CopyFilteredRows", "Phase column index (L) = " & phaseColIndex

    ' --- Filter rows in memory into potentially oversized outputData ---
    Module_Dashboard.DebugLog "CopyFilteredRows", "Filtering " & UBound(arrData, 1) & " rows... KeepArchived=" & keepArchived & ", phaseColIndex=" & phaseColIndex

    outputRowCount = 0
    For r = LBound(arrData, 1) To UBound(arrData, 1) ' Loop through source data array rows
        ' Get phase from the correct numeric index, trim, convert to string defensively
        phaseValue = Trim$(CStr(arrData(r, phaseColIndex)))

        ' Sample the first 10 rows for debugging filter logic (Commented out 04/20/2025 - Too verbose)
        ' If r <= 10 Then _
        '     Module_Dashboard.DebugLog "FilterSample", "r=" & r & ", phase='" & phaseValue & "', IsPhaseArchived=" & IsPhaseArchived(phaseValue) & ", IsPhaseActive=" & IsPhaseActive(phaseValue)

        ' Check if the row's archived status matches the filter requirement
        ' If keepArchived=True, we want rows where IsPhaseArchived is True
        ' If keepArchived=False, we want rows where IsPhaseArchived is False (i.e., Active or Blank)
        If IsPhaseArchived(phaseValue) = keepArchived Then
            outputRowCount = outputRowCount + 1 ' Increment count of rows that match filter
            ' Copy matching row data from source array to output array
            For i = 1 To numCols
                outputData(outputRowCount, i) = arrData(r, i)
            Next i
        End If
    Next r

    Module_Dashboard.DebugLog "CopyFilteredRows", "Filtering complete. Matching rows = " & outputRowCount
    ' *** Sanity-check: if nothing was copied but the source had data, log a warning ***
    If outputRowCount = 0 And UBound(arrData, 1) >= 1 Then
        Module_Dashboard.DebugLog "CopyFilteredRows", "WARNING – zero rows met the criteria; source had " & UBound(arrData, 1) & " rows."
    End If

    ' <<< ACTION: SET COUNT PROPERTY BASED ON FINAL outputRowCount >>>
    ' After the loop, store the final count for this view type in the corresponding property
    If keepArchived Then
        ArchiveRecords = outputRowCount ' *** FIXED: Removed Me. *** SET: Store Archive count via Property Let
        Module_Dashboard.DebugLog "CopyFilteredRows", "Set ArchiveRecords Property: " & outputRowCount
    Else ' keepArchived was False, so this was the Active view
        ActiveRecords = outputRowCount ' *** FIXED: Removed Me. *** SET: Store Active count via Property Let
        Module_Dashboard.DebugLog "CopyFilteredRows", "Set ActiveRecords Property: " & outputRowCount
    End If
    ' <<< END: SET COUNT PROPERTY >>>


    ' --- Write results to Target Sheet ---
    Module_Dashboard.DebugLog "CopyFilteredRows", "Writing results to target sheet '" & wsTgt.Name & "'..."
    wsTgt.Cells.ClearContents ' Clear target completely first
    Module_Dashboard.DebugLog "CopyFilteredRows", "Cleared target sheet."
    wsSrc.Range("A1:" & lastColLetter & "3").Copy wsTgt.Range("A1") ' Copy A1:N3 (Title/Control/Headers)
    Module_Dashboard.DebugLog "CopyFilteredRows", "Copied headers A1:" & lastColLetter & "3."

    '----- WRITE THE FILTERED ROWS -----
    If outputRowCount > 0 Then

        ' --- FIX for ReDim Preserve 32-bit Issue / General Best Practice: Copy to correctly sized array ---
        Module_Dashboard.DebugLog "CopyFilteredRows", "Preparing finalOutputData array sized " & outputRowCount & "x" & numCols
        ReDim finalOutputData(1 To outputRowCount, 1 To numCols) ' Size exactly right
        Module_Dashboard.DebugLog "CopyFilteredRows", "Copying " & outputRowCount & " rows from temp outputData to finalOutputData..."
        For r = 1 To outputRowCount
            For i = 1 To numCols
                finalOutputData(r, i) = outputData(r, i) ' Copy relevant rows/cols
            Next i
        Next r
        Module_Dashboard.DebugLog "CopyFilteredRows", "Finished copying to finalOutputData."
        Erase outputData ' Release memory of potentially oversized temporary array
        ' --- End Fix ---

        ' --- Pre-Write Checks (Optional but good for debugging hard-to-find write errors) ---
        Dim isProtected As Boolean
        Dim isMerged As Boolean
        On Error Resume Next ' Check properties safely
        isProtected = wsTgt.ProtectContents
        isMerged = wsTgt.Range("A4").MergeCells ' Check A4 specifically for merge (common cause of array write failure)
        On Error GoTo CopyFilteredRows_Error ' Restore main handler
        Module_Dashboard.DebugLog "CopyFilteredRows", "PRE-WRITE CHECK: wsTgt.ProtectContents = " & isProtected
        Module_Dashboard.DebugLog "CopyFilteredRows", "PRE-WRITE CHECK: wsTgt.Range(""A4"").MergeCells = " & isMerged
        ' --- End Pre-Write Checks ---

        Module_Dashboard.DebugLog "CopyFilteredRows", "Attempting write finalOutputData to " & wsTgt.Name & ".Range(""A4"")..."
        On Error GoTo WriteErr ' *** Trap errors JUST for the array write operation ***
        wsTgt.Range("A4").Resize(outputRowCount, numCols).value = finalOutputData ' <<< WRITE finalOutputData >>>
        On Error GoTo CopyFilteredRows_Error ' *** Restore main handler AFTER successful write ***

        ' If we get here, write was successful
        Module_Dashboard.DebugLog "CopyFilteredRows", "Wrote " & outputRowCount & " rows to '" & wsTgt.Name & "'."
        GoTo AfterWrite ' Skip the specific write error handler block

WriteErr: ' *** Specific handler for array write error ***
        Module_Dashboard.DebugLog "CopyFilteredRows", "*** WRITE FAILED *** Err " & Err.Number & ": " & Err.Description
        MsgBox "CopyFilteredRows failed writing data to '" & wsTgt.Name & "'" & vbCrLf & _
               "Check for merged cells, protection, or other issues." & vbCrLf & _
               "Err " & Err.Number & ": " & Err.Description, vbCritical, "Data Write Error"
        On Error GoTo CopyFilteredRows_Error ' Restore main handler before exiting
        Erase arrData: If IsArray(finalOutputData) Then Erase finalOutputData ' Clean up arrays
        Exit Sub ' Exit after write failure

AfterWrite: ' Label to jump to after successful write
        ' Any code needed immediately after successful write could go here (e.g., specific formatting)

    Else ' outputRowCount = 0
        Module_Dashboard.DebugLog "CopyFilteredRows", "No rows matched filter criteria for '" & wsTgt.Name & "'. Clearing data area below headers."
        ' Ensure data area is clear if no rows written
        wsTgt.Range("A4:" & lastColLetter & wsTgt.rows.Count).ClearContents
    End If
    ' --- End of write block ---

    ' <<< ACTION: REMOVED the block that wrote the local count to J2 >>>
    ' The final counts will be written later by the calling module after reading properties.

    ' Clean up memory
    Erase arrData
    If IsArray(finalOutputData) Then Erase finalOutputData ' Erase the final array if it was created
    Module_Dashboard.DebugLog "CopyFilteredRows", "EXIT (Normal). Time: " & Format(Timer - t1, "0.00") & "s"
    Exit Sub ' Normal Exit

CopyFilteredRows_Error:
    If Err.Number <> 0 Then ' Log only real errors
        Module_Dashboard.DebugLog "CopyFilteredRows", "!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!"
        Module_Dashboard.DebugLog "CopyFilteredRows", "ERROR Handler! Err=" & Err.Number & ": " & Err.Description & " near line " & Erl
        Module_Dashboard.DebugLog "CopyFilteredRows", "!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!"
    End If
    ' Clean up memory even on error
    If IsArray(arrData) Then Erase arrData
    If IsArray(outputData) Then Erase outputData ' Original temp array might still exist if error was early
    If IsArray(finalOutputData) Then Erase finalOutputData ' Final array
    ' Let error propagate back to the caller's error handler if not handled by WriteErr
End Sub

Private Function IsPhaseActive(ph As String) As Boolean
    '--------------------------------------------------
    ' Purpose: Checks if a given phase string exists in the pipe-delimited ACTIVE_PHASES constant.
    '          Used to identify phases that should exclusively be on the Active sheet.
    ' Logic:   Performs a case-insensitive check for "|PHASE|" within the constant.
    ' Returns: True if found, False otherwise (including for blank strings).
    ' Called By: IsPhaseArchived
    '--------------------------------------------------
    ph = UCase$(Trim$(ph)) ' Ensure uppercase and trimmed for comparison
    If Len(ph) = 0 Then
         IsPhaseActive = False ' Blank is NOT explicitly Active
    Else
         ' Check if "|PHASE|" exists within "|PHASE1|PHASE2|..." using case-insensitive comparison
         IsPhaseActive = (InStr(1, ACTIVE_PHASES, "|" & ph & "|", vbTextCompare) > 0)
    End If
End Function

Private Function IsPhaseArchived(ByVal phase As String) As Boolean
    '--------------------------------------------------
    ' Purpose: Determines if a phase should be classified as "Archived".
    ' Logic:   A phase is considered Archived if it is NOT blank AND it is NOT explicitly Active
    '          (as determined by the IsPhaseActive function checking the ACTIVE_PHASES constant).
    '          Blank phases are treated as NOT Archived (effectively Active by default in the filtering logic).
    ' Returns: True if the phase is non-blank and not in ACTIVE_PHASES, False otherwise.
    ' Called By: CopyFilteredRows
    '--------------------------------------------------
    phase = Trim$(phase) ' Trim whitespace
    If Len(phase) = 0 Then
        IsPhaseArchived = False ' Blank is considered NOT Archived (will be kept when keepArchived=False)
    Else
        ' If it's not explicitly Active, it belongs in the Archive.
        IsPhaseArchived = Not IsPhaseActive(phase)
    End If
End Function


'--------------------------------------------------
' SECTION 8: SHEET MANAGEMENT UTILITIES
'--------------------------------------------------
' Contains helper functions for checking sheet existence and creating/preparing sheets.

Private Function SheetExists(sName As String) As Boolean
    ' Purpose: Checks if a sheet (of any type) with the given name exists in ThisWorkbook.
    ' Returns: True if a sheet with the specified name exists, False otherwise.
    ' Called By: RefreshActiveView, RefreshArchiveView (Defensive Checks), GetOrCreateSheet
    Dim Sh As Object ' Use Object to check any sheet type (Worksheet, Chart, etc.)
    On Error Resume Next ' Prevent error if sheet doesn't exist, allowing Sh to remain Nothing
    Set Sh = ThisWorkbook.Sheets(sName)
    SheetExists = Not Sh Is Nothing ' If Sh is Nothing, sheet doesn't exist
    On Error GoTo 0 ' Restore default error handling
    Set Sh = Nothing ' Release object
End Function

Private Function GetOrCreateSheet(sName As String, sTitle As String, _
                                bannerColor As Long) As Worksheet
    ' Purpose: Retrieves a worksheet by name if it exists, or creates, names,
    '          and formats a new worksheet if it doesn't. Handles potential conflicts
    '          with non-worksheet objects of the same name and workbook structure protection.
    ' Returns: Worksheet object if successful, Nothing otherwise.
    ' Called By: RefreshActiveView, RefreshArchiveView
    Dim ws As Worksheet
    Dim wbStructureWasLocked As Boolean
    Dim t1 As Double: t1 = Timer

    Module_Dashboard.DebugLog "GetOrCreateSheet", "ENTER. sName='" & sName & "'"
    On Error GoTo GetOrCreateSheet_Error ' Use specific handler for this function

    ' --- Check if worksheet exists ---
    Module_Dashboard.DebugLog "GetOrCreateSheet", "Checking if worksheet exists..."
    Set ws = Nothing ' Ensure ws is Nothing initially
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(sName) ' Try to get as Worksheet specifically
    If Err.Number <> 0 Then Err.Clear ' Clear error if not found or wrong type (e.g., Chart sheet)
    On Error GoTo GetOrCreateSheet_Error ' Restore handler

    If ws Is Nothing Then
        ' Worksheet does not exist. Check if *any* sheet type exists with this name.
        Dim anySheet As Object
        Set anySheet = Nothing
        On Error Resume Next
        Set anySheet = ThisWorkbook.Sheets(sName) ' Check for Charts, etc.
        On Error GoTo GetOrCreateSheet_Error
        If Not anySheet Is Nothing Then
             Module_Dashboard.DebugLog "GetOrCreateSheet", "WARNING: A non-worksheet object named '" & sName & "' exists (Type: " & TypeName(anySheet) & "). Renaming it."
             On Error Resume Next ' Handle error renaming
             anySheet.Name = sName & "_Old_" & Format(Now, "hhmmss") ' Append timestamp to rename
             If Err.Number <> 0 Then
                 Module_Dashboard.DebugLog "GetOrCreateSheet", "ERROR: Failed to rename existing non-worksheet '" & sName & "'. Aborting. Err=" & Err.Number & ": " & Err.Description
                 Set GetOrCreateSheet = Nothing: Exit Function ' Cannot proceed
             End If
             Err.Clear
             On Error GoTo GetOrCreateSheet_Error
             Set anySheet = Nothing ' Release object
        End If

        ' Proceed with worksheet creation
        Module_Dashboard.DebugLog "GetOrCreateSheet", "Worksheet '" & sName & "' not found or was renamed. Attempting creation..."
        ' --- Need to temporarily unprotect workbook structure to add sheet ---
        wbStructureWasLocked = ThisWorkbook.ProtectStructure
        Module_Dashboard.DebugLog "GetOrCreateSheet", "Workbook Structure locked = " & wbStructureWasLocked
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
        ' Add sheet after the last existing sheet
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
        On Error Resume Next ' Handle error during Name (e.g., invalid characters, duplicate name somehow)
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
         ' Optional: Clear contents if sheet already exists? Typically handled by CopyFilteredRows.
         ' Module_Dashboard.DebugLog "GetOrCreateSheet", "Clearing existing sheet contents..."
         ' ws.Cells.Clear ' Uncomment if desired behavior is to always clear on retrieval
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
    Dim lastColLetter As String: lastColLetter = Module_Dashboard.DB_COL_COMMENTS ' Should be "N" (from Module_Dashboard)
    On Error Resume Next ' Handle error if sheet is protected or other issues applying format
    With ws.Range("A1:" & lastColLetter & "1") ' Use Constant for last column
        If .MergeCells Then .UnMerge ' Ensure not already merged incorrectly
        .ClearContents
        .Merge ' Merge cells A1 to N1
        .value = sTitle ' Set the title text
        .Font.Bold = True: .Font.Size = 16 ' Format font
        .Interior.Color = bannerColor ' Set background color passed as argument
        .Font.Color = vbWhite ' White text for contrast
        .HorizontalAlignment = xlCenter ' Center align text
        .VerticalAlignment = xlCenter ' Center align text
        .RowHeight = 32 ' Set standard title row height
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
' SECTION 9: FORMATTING & UI HELPERS
'--------------------------------------------------
' Contains routines that apply formatting, create UI elements (like buttons),
' and manage sheet protection for the view sheets.

Private Sub ApplyViewFormatting(ws As Worksheet, viewTag As String) ' viewTag is "Active" or "Archive" for logging/context
    ' Purpose: Applies standard formatting (column widths, number formats, conditional formatting),
    '          sets row heights, applies freeze panes, adds navigation buttons, and protects the sheet.
    ' Called By: RefreshActiveView, RefreshArchiveView
    Dim lastRow As Long
    Dim lastColLetter As String: lastColLetter = Module_Dashboard.DB_COL_COMMENTS ' Should be "N"
    Dim phaseColLetter As String: phaseColLetter = Module_Dashboard.DB_COL_PHASE ' Should be "L"
    Dim wsSrc As Worksheet ' To copy column widths from
    Dim t1 As Double: t1 = Timer

    On Error GoTo ApplyViewFormatting_Error ' Use specific handler for this sub
    Module_Dashboard.DebugLog "ApplyViewFormatting", "ENTER. TargetSheet='" & ws.Name & "', ViewTag='" & viewTag & "'"

    If ws Is Nothing Then
        Module_Dashboard.DebugLog "ApplyViewFormatting", "ERROR: ws object Is Nothing on entry. Exiting."
        Exit Sub
    End If

    Application.ScreenUpdating = False ' Keep off during formatting to improve speed and reduce flicker

    ' --- Unprotect first ---
    Module_Dashboard.DebugLog "ApplyViewFormatting", "Unprotecting sheet..."
    On Error Resume Next ' Handle error if already unprotected or password wrong
    ws.Unprotect Password:=Module_Dashboard.PW_WORKBOOK ' Use password constant from Module_Dashboard
    If Err.Number <> 0 Then Module_Dashboard.DebugLog "ApplyViewFormatting", "Warning: Error unprotecting sheet. Err=" & Err.Number: Err.Clear
    On Error GoTo ApplyViewFormatting_Error ' Restore handler

    ' --- Optional: Set consistent data row height ---
    ' Calculate lastRow here to apply height to data rows
    lastRow = ws.Cells(ws.rows.Count, "A").End(xlUp).Row ' Calculate last data row based on Col A
    If lastRow >= 4 Then ' Data starts row 4
        ws.rows("4:" & lastRow).RowHeight = 18 ' Example: Set data rows to 18 points
    End If
    If Err.Number <> 0 Then Module_Dashboard.DebugLog "ApplyViewFormatting", "Warning: Error setting row heights. Err=" & Err.Number: Err.Clear
    On Error GoTo ApplyViewFormatting_Error ' Restore main error handler
    Module_Dashboard.DebugLog "ApplyViewFormatting", "Applied consistent row heights for Rows 1, 2" & IIf(lastRow >= 4, " and Data Rows.", ".")
    ' --- End Row Height Fix ---

    ' --- Apply Column Widths (Copy from Source Dashboard) ---
    Module_Dashboard.DebugLog "ApplyViewFormatting", "Applying column widths..."
    On Error Resume Next
    Set wsSrc = ThisWorkbook.Worksheets(Module_Dashboard.DASHBOARD_SHEET_NAME) ' Get source dashboard sheet
    If Err.Number <> 0 Or wsSrc Is Nothing Then
        Module_Dashboard.DebugLog "ApplyViewFormatting", "Warning - Could not get source dashboard sheet '" & Module_Dashboard.DASHBOARD_SHEET_NAME & "' to copy widths. Err=" & Err.Number
        Err.Clear
        ' Fallback: AutoFit A:N if source widths unavailable
        Module_Dashboard.DebugLog "ApplyViewFormatting", "Applying AutoFit as fallback..."
        ws.Columns("A:" & lastColLetter).AutoFit
    Else
        ' Copy widths from source dashboard A1:N1 range
        wsSrc.Range("A1:" & lastColLetter & "1").Copy
        ws.Range("A1").PasteSpecial xlPasteColumnWidths
        Application.CutCopyMode = False ' Clear clipboard marquee
        Module_Dashboard.DebugLog "ApplyViewFormatting", "Copied column widths from '" & wsSrc.Name & "'."
    End If
    Set wsSrc = Nothing ' Release source sheet object
    On Error GoTo ApplyViewFormatting_Error ' Restore handler

    ' --- Apply Number Formatting ---
    Module_Dashboard.DebugLog "ApplyViewFormatting", "Applying number formats..."
    ' lastRow was calculated earlier for row height fix
    Module_Dashboard.DebugLog "ApplyViewFormatting", "lastRow for formatting = " & lastRow
    ' Data starts at Row 4 on these views
    If lastRow >= 4 Then
        On Error Resume Next ' Handle errors applying formats (e.g., invalid data)
        ' Apply formats using column constants from Module_Dashboard where available
        ws.Range(Module_Dashboard.DB_COL_AMOUNT & "4:" & Module_Dashboard.DB_COL_AMOUNT & lastRow).NumberFormat = "$#,##0.00"  ' Amount (D)
        ws.Range(Module_Dashboard.DB_COL_DOC_DATE & "4:" & Module_Dashboard.DB_COL_DOC_DATE & lastRow).NumberFormat = "mm/dd/yyyy" ' Document Date (E)
        ws.Range(Module_Dashboard.DB_COL_FIRST_PULL & "4:" & Module_Dashboard.DB_COL_FIRST_PULL & lastRow).NumberFormat = "mm/dd/yyyy" ' First Date Pulled (F)
        ws.Range(Module_Dashboard.DB_COL_PULL_COUNT & "4:" & Module_Dashboard.DB_COL_PULL_COUNT & lastRow).NumberFormat = "0"           ' Pull Count (I)
        ws.Range(Module_Dashboard.DB_COL_LASTCONTACT & "4:" & Module_Dashboard.DB_COL_LASTCONTACT & lastRow).NumberFormat = "mm/dd/yyyy" ' Last Contact (M)
        If Err.Number <> 0 Then Module_Dashboard.DebugLog "ApplyViewFormatting", "Warning: Error applying number formats. Err=" & Err.Number: Err.Clear
        On Error GoTo ApplyViewFormatting_Error ' Restore handler
    Else
         Module_Dashboard.DebugLog "ApplyViewFormatting", "Skipping number formats (lastRow < 4)."
    End If

    ' --- Apply Conditional Formatting ---
    If lastRow >= 4 Then ' Data starts row 4
         Module_Dashboard.DebugLog "ApplyViewFormatting", "Applying conditional formatting..."
         ' Call helper routines from Module_Dashboard to apply standard CF rules
         Module_Dashboard.DebugLog "ApplyViewFormatting", "Calling ApplyColorFormatting..."
         Module_Dashboard.ApplyColorFormatting ws, 4 ' Start formatting data from row 4
         Module_Dashboard.DebugLog "ApplyViewFormatting", "Calling ApplyWorkflowLocationFormatting..."
         Module_Dashboard.ApplyWorkflowLocationFormatting ws, 4 ' Start formatting data from row 4
         Module_Dashboard.DebugLog "ApplyViewFormatting", "Finished applying conditional formatting."
    Else
         Module_Dashboard.DebugLog "ApplyViewFormatting", "Skipping conditional formatting (lastRow < 4)."
    End If

' --- Apply Data Validation (Phase Column) ---
    ' Ensure the dropdown list for Phase is present even on read-only sheets for consistency/clarity
    Module_Dashboard.DebugLog "ApplyViewFormatting", "Applying Phase data validation to column " & Module_Dashboard.DB_COL_PHASE & "..."
    ' *** FIXED: Call the sub from modUtilities where it now resides ***
    Call modUtilities.ApplyPhaseValidationToListColumn(ws, Module_Dashboard.DB_COL_PHASE, 4) ' Apply to Col L, starting Row 4
    Module_Dashboard.DebugLog "ApplyViewFormatting", "Applied Phase data validation."

    ' --- Protection (Make read-only) ---
    Module_Dashboard.DebugLog "ApplyViewFormatting", "Locking all cells..."
    On Error Resume Next ' Handle error if sheet is protected
    ws.Cells.Locked = True ' Lock all cells on the view sheets first
    If Err.Number <> 0 Then Module_Dashboard.DebugLog "ApplyViewFormatting", "Warning: Error locking all cells. Err=" & Err.Number: Err.Clear
    On Error GoTo ApplyViewFormatting_Error ' Restore handler

    ' --- Apply Freeze Panes ---
    Module_Dashboard.DebugLog "ApplyViewFormatting", "Applying freeze panes..."
    On Error Resume Next ' Handle errors if sheet not active or already frozen/unfrozen
    ws.Activate ' Must activate to set freeze panes
    ActiveWindow.FreezePanes = False ' Unfreeze first to ensure correct state
    ws.Range("A4").Select           ' Select cell below header/control rows (A1:N3)
    ActiveWindow.FreezePanes = True   ' Freeze Rows 1-3
    ws.Range("A1").Select ' Select A1 after freezing for better user experience
    If Err.Number <> 0 Then Module_Dashboard.DebugLog "ApplyViewFormatting", "Warning: Error applying freeze panes. Err=" & Err.Number: Err.Clear
    On Error GoTo ApplyViewFormatting_Error ' Restore handler
    Module_Dashboard.DebugLog "ApplyViewFormatting", "Applied freeze panes."

    ' --- Apply EXACT Dashboard Formatting Clone (Row 1 Title, Row 2 Controls, Buttons, Counts) ---
    ' Calls the central cloning function from modFormatting to ensure pixel-perfect replication.
    Module_Dashboard.DebugLog "ApplyViewFormatting", "Calling modFormatting.ExactlyCloneDashboardFormatting..."
    modFormatting.ExactlyCloneDashboardFormatting ws, viewTag ' Use the exact cloning function
    Module_Dashboard.DebugLog "ApplyViewFormatting", "Returned from ExactlyCloneDashboardFormatting."

     ' --- Final Protection (Specific to Active/Archive Views) ---
     ' This protection is applied AFTER all formatting and UI elements are set.
     Module_Dashboard.DebugLog "ApplyViewFormatting", "Applying final sheet protection..."
     On Error Resume Next ' Handle protection errors
     ws.Protect Password:=Module_Dashboard.PW_WORKBOOK, DrawingObjects:=True, Contents:=True, Scenarios:=True, _
         AllowFormattingCells:=False, AllowFormattingColumns:=False, AllowFormattingRows:=False, _
         AllowInsertingColumns:=False, AllowInsertingRows:=False, AllowInsertingHyperlinks:=False, _
         AllowDeletingColumns:=False, AllowDeletingRows:=False, AllowSorting:=False, AllowFiltering:=False, _
         AllowUsingPivotTables:=False ' Apply restrictive protection for read-only views
     If Err.Number <> 0 Then Module_Dashboard.DebugLog "ApplyViewFormatting", "Warning: Error applying sheet protection. Err=" & Err.Number: Err.Clear
     On Error GoTo ApplyViewFormatting_Error ' Restore handler
     Module_Dashboard.DebugLog "ApplyViewFormatting", "Applied sheet protection (Read-Only)."

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

Public Sub AddNavigationButtons(ws As Worksheet) ' Made Public as it's called by ApplyViewFormatting (could be Private if ApplyViewFormatting is always the caller)
    ' Purpose: Creates the standard set of navigation and refresh buttons in Row 2 (C,D,F,G,H)
    '          and adds the "Last Refresh" timestamp in N2.
    '          REMOVED all code writing counts/placeholders to J2:L2. Count writing is handled
    '          by a central routine (e.g., modUtilities.UpdateAllViewCounts) that reads the properties.
    ' Called By: ApplyViewFormatting (for Active/Archive sheets), Module_Dashboard.RefreshDashboard (for main Dashboard sheet)

    Const TOPROW As Long = 2
    Const BTN_H As Double = 24      ' Standard button height (used by ModernButton)
    ' Button widths are defined in the btnDefs array below
    Dim btnDefs As Variant, def As Variant, shp As Shape, target As Range, i As Long
    Dim t1 As Double: t1 = Timer

    ' --- Define each button: { Target Cell Address, Caption, Macro Name, Width } ---
    ' Width = 0 means autofit to cell width (minus padding) - handled by ModernButton logic
    ' Width > 0 means fixed width (used for the smaller nav buttons)
    btnDefs = Array( _
      Array("C2", "Standard Refresh", "Button_RefreshDashboard_SaveAndRestoreEdits", 0), _
      Array("D2", "Preserve UserEdits", "Button_RefreshDashboard_PreserveUserEdits", 0), _
      Array("F2", "All Items", "modArchival.btnViewAll", 65), _
      Array("G2", "Active", "modArchival.btnViewActive", 65), _
      Array("H2", "Archive", "modArchival.btnViewArchive", 65) _
    )

    On Error GoTo AddNav_Error
    Module_Dashboard.DebugLog "AddNavigationButtons", "ENTER for sheet: '" & ws.Name & "' - Layout C2/D2/F2/G2/H2"

    If ws Is Nothing Then Exit Sub

    '----- preparation: Unprotect, Clear Row 2 Shapes/Content/Merges -----
    On Error Resume Next
    ws.Unprotect Password:=Module_Dashboard.PW_WORKBOOK ' Unprotect sheet to modify shapes/content
    Dim deleteCount As Long: deleteCount = 0
    For Each shp In ws.Shapes ' Delete ANY shape anchored anywhere in Row 2
        If shp.TopLeftCell.Row = TOPROW Then
            ' Be specific about deleting only buttons created by this process if possible
            ' Using name patterns or checking Type can help avoid deleting other shapes.
            If shp.Name Like "btn_*" Or shp.Name Like "nav*" Or shp.Type = msoFormControl Then ' Example patterns
                shp.Delete ' Delete existing buttons in row 2
                deleteCount = deleteCount + 1
            End If
        End If
    Next shp
    ws.Range("B" & TOPROW & ":N" & TOPROW).ClearContents ' Clear all cell content from B2:N2 (Leaves A2 intact if needed)
    ws.Range("B" & TOPROW & ":N" & TOPROW).UnMerge       ' Unmerge all relevant cells just in case
    If Err.Number <> 0 Then Module_Dashboard.DebugLog "AddNavigationButtons", "Warning during clear/unmerge: Err=" & Err.Number: Err.Clear
    On Error GoTo AddNav_Error ' Restore error handler
    Module_Dashboard.DebugLog "AddNavigationButtons", "Cleared Row " & TOPROW & ". Deleted " & deleteCount & " shapes."

    '----- create the five buttons ------------------------------------
    Module_Dashboard.DebugLog "AddNavigationButtons", "Creating buttons..."
    For i = LBound(btnDefs) To UBound(btnDefs) ' Loop through button definitions array
        def = btnDefs(i) ' Get definition for current button
        Set target = Nothing ' Reset target object
        On Error Resume Next
        Set target = ws.Range(def(0)) ' Get target Cell object (e.g., Range("C2"))
        On Error GoTo AddNav_Error

        If target Is Nothing Then
            Module_Dashboard.DebugLog "AddNavigationButtons", "ERROR: Invalid target cell '" & def(0) & "'. Skipping button creation."
        Else
            Set shp = Nothing ' Reset shape object
            ' Call ModernButton factory function (in Module_Dashboard) to create the styled button shape
            Set shp = Module_Dashboard.ModernButton(ws, target, def(1), def(2), def(3)) ' Pass definition array elements: ws, TargetCell, Caption, Macro, Width

            If Not shp Is Nothing Then ' Check if button creation succeeded

                ' --- Adjust Height for Nav Buttons (Optional visual tweak) ---
                If def(0) = "F2" Or def(0) = "G2" Or def(0) = "H2" Then
                    shp.Height = 18 ' Use shorter height for F/G/H buttons for visual distinction
                    Module_Dashboard.DebugLog "AddNavigationButtons", "Adjusted height to 18 for: " & def(1)
                End If
                ' Note: Width is handled by ModernButton based on def(3) parameter (0=autofit, >0=fixed)

                ' --- Final Naming and Centering ---
                On Error Resume Next ' Handle errors modifying shape properties
                shp.Name = "btn_" & Replace(def(1), " ", "_") ' Set final name (e.g., btn_Standard_Refresh)
                ' Recalculate Left/Top for precise centering within the target cell after potential height/width adjustments
                shp.Left = target.Left + (target.Width - shp.Width) / 2
                shp.Top = target.Top + (target.Height - shp.Height) / 2
                If Err.Number <> 0 Then Module_Dashboard.DebugLog "AddNavigationButtons", "Warning: Error centering/naming '" & shp.Name & "'. Err=" & Err.Number: Err.Clear
                On Error GoTo AddNav_Error ' Restore handler
            Else
                Module_Dashboard.DebugLog "AddNavigationButtons", "ERROR: ModernButton failed for '" & def(1) & "'."
            End If
        End If
    Next i
    Module_Dashboard.DebugLog "AddNavigationButtons", "Finished creating buttons."

    '----- timestamp ONLY -----------------------------
    ' <<< ACTION: REMOVED the With block that wrote counts/placeholders to J2:L2 >>>
    ' Count writing is now handled by a separate routine that reads the properties.

    Module_Dashboard.DebugLog "AddNavigationButtons", "Setting timestamp in N2..."
    On Error Resume Next
    With ws.Range("N" & TOPROW) ' N2 - Timestamp ONLY
        .value = "Refreshed: " & Format$(Now(), "mm/dd hh:nn AM/PM") ' Use your preferred format
        .Font.Size = 9
        .Font.Italic = True ' Italic is fine for the timestamp aesthetic
        .HorizontalAlignment = xlLeft ' Align left within N2
        .VerticalAlignment = xlCenter
    End With
    If Err.Number <> 0 Then Module_Dashboard.DebugLog "AddNavigationButtons", "Warning: Error setting timestamp text. Err=" & Err.Number: Err.Clear
    On Error GoTo AddNav_Error ' Restore handler

    '----- re-protect REMOVED (04/20/2025) -----------------------------
    ' Protection is now handled by the calling function (ApplyViewFormatting)
    ' AFTER all formatting and UI elements (including counts) are complete.
    ' This prevents errors where cells were locked before counts could be written.
    Module_Dashboard.DebugLog "AddNavigationButtons", "Skipping re-protection within this sub."


    Module_Dashboard.DebugLog "AddNavigationButtons", "EXIT (Normal - Buttons/Timestamp Only). Time: " & Format(Timer - t1, "0.00") & "s"
    Exit Sub

AddNav_Error: ' Error Handler for this subroutine
    Module_Dashboard.DebugLog "AddNavigationButtons", "!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!"
    Module_Dashboard.DebugLog "AddNavigationButtons", "ERROR Handler! Sheet='" & IIf(ws Is Nothing, "UNKNOWN", ws.Name) & "'. Err=" & Err.Number & ": " & Err.Description & " near line " & Erl
    Module_Dashboard.DebugLog "AddNavigationButtons", "!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!"
    Set shp = Nothing: Set target = Nothing ' Clean up objects
    ' Consider whether to attempt re-protection on error if sheet was unprotected
End Sub
