Attribute VB_Name = "Module_Dashboard"
Option Explicit

' --- Constants ---
Private Const DASHBOARD_SHEET_NAME As String = "SQRCT Dashboard"
Private Const USEREDITS_SHEET_NAME As String = "UserEdits"
Private Const USEREDITSLOG_SHEET_NAME As String = "UserEditsLog"
Private Const MASTER_QUOTES_FINAL_SOURCE As String = "MasterQuotes_Final" ' Name of the PQ query/table

' UserEdits Columns (Adjusted for Email removal)
Private Const UE_COL_DOCNUM As String = "A"
Private Const UE_COL_PHASE As String = "B"
Private Const UE_COL_LASTCONTACT As String = "C"
' Private Const UE_COL_EMAIL As String = "D" ' REMOVED - Column Shift Required Below
Private Const UE_COL_COMMENTS As String = "D" ' Shifted from E
Private Const UE_COL_SOURCE As String = "E"   ' Shifted from F
Private Const UE_COL_TIMESTAMP As String = "F" ' Shifted from G

' Dashboard Columns (Editable - Adjusted for Email removal)
Private Const DB_COL_PHASE As String = "K"
Private Const DB_COL_LASTCONTACT As String = "L"
' Private Const DB_COL_EMAIL As String = "M" ' REMOVED - Column Shift Required Below
Private Const DB_COL_COMMENTS As String = "M" ' Shifted from N
' --- End Constants ---


'===============================================================================
' MODULE_DASHBOARD
' Contains functions for managing the SQRCT Dashboard, including refresh operations,
' user edits management, and UI interactions.
' Refactored for performance (Dictionary lookup), security (no passwords), and robustness (constants).
' Protection logic revised: Applied after modifications in relevant subs.
'===============================================================================

'===============================================================================
' BUTTON FUNCTIONS: These can be assigned to buttons in the UI
'===============================================================================

' Standard workflow - saves dashboard edits first, then restores after refresh
Public Sub Button_RefreshDashboard_SaveAndRestoreEdits()
    RefreshDashboard_TwoWaySync
End Sub

' Special workflow - only updates dashboard from UserEdits without saving current edits
Public Sub Button_RefreshDashboard_PreserveUserEdits()
    RefreshDashboard_OneWayFromUserEdits
End Sub

'===============================================================================
' MAIN FUNCTIONS: with clear names indicating data flow direction
'===============================================================================

' Standard two-way sync: Dashboard -> UserEdits -> Refresh -> UserEdits -> Dashboard
Public Sub RefreshDashboard_TwoWaySync()
    Call RefreshDashboard(PreserveUserEdits:=False)
End Sub

' One-way sync: Refresh -> UserEdits -> Dashboard (preserves manual UserEdits)
Public Sub RefreshDashboard_OneWayFromUserEdits()
    Call RefreshDashboard(PreserveUserEdits:=True)
End Sub

'===============================================================================
' USER EDITS LOGGING: Tracks operations performed on UserEdits
'===============================================================================
Public Sub LogUserEditsOperation(message As String)
    Dim wsLog As Worksheet
    Dim lastRow As Long

    On Error Resume Next
    Set wsLog = ThisWorkbook.Sheets(USEREDITSLOG_SHEET_NAME) ' Use Constant

    If wsLog Is Nothing Then
        Set wsLog = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        wsLog.Name = USEREDITSLOG_SHEET_NAME ' Use Constant
        wsLog.Range("A1:C1").Value = Array("Timestamp", "Workbook", "Operation")
        wsLog.Range("A1:C1").Font.Bold = True
        wsLog.Visible = xlSheetHidden
    End If
    On Error GoTo 0 ' Restore default error handling before proceeding

    lastRow = wsLog.Cells(wsLog.Rows.Count, "A").End(xlUp).Row
    If lastRow < 1 Then lastRow = 1

    wsLog.Cells(lastRow + 1, "A").Value = Format(Now, "yyyy-mm-dd hh:mm:ss")
    wsLog.Cells(lastRow + 1, "B").Value = GetWorkbookIdentity()  ' Using Module_Identity function
    wsLog.Cells(lastRow + 1, "C").Value = message

    On Error GoTo 0
End Sub

'===============================================================================
' USER EDITS BACKUP: Creates a backup of the UserEdits sheet
'===============================================================================
Public Function CreateUserEditsBackup(Optional backupSuffix As String = "") As Boolean
    On Error GoTo ErrorHandler

    Dim wsEdits As Worksheet
    Dim wsBackup As Worksheet
    Dim backupName As String

    ' Set backup name
    If backupSuffix = "" Then
        backupName = USEREDITS_SHEET_NAME & "_Backup_" & Format(Now, "yyyymmdd") ' Use Constant
    Else
        backupName = USEREDITS_SHEET_NAME & "_Backup_" & backupSuffix ' Use Constant
    End If

    ' Create backup only if UserEdits exists and has data
    On Error Resume Next
    Set wsEdits = ThisWorkbook.Sheets(USEREDITS_SHEET_NAME) ' Use Constant
    On Error GoTo ErrorHandler

    If wsEdits Is Nothing Then
        CreateUserEditsBackup = False
        Exit Function
    End If

    ' Check if the backup sheet already exists
    On Error Resume Next
    Set wsBackup = ThisWorkbook.Sheets(backupName)
    If wsBackup Is Nothing Then
        Set wsBackup = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        wsBackup.Name = backupName
    Else
        wsBackup.Cells.Clear
    End If
    On Error GoTo ErrorHandler

    ' Copy data
    wsEdits.UsedRange.Copy wsBackup.Range("A1")

    ' Hide backup sheet
    wsBackup.Visible = xlSheetHidden

    ' Log the backup creation
    LogUserEditsOperation "Created UserEdits backup: " & backupName

    CreateUserEditsBackup = True
    Exit Function

ErrorHandler:
    Debug.Print "Error creating backup: " & Err.Description
    LogUserEditsOperation "ERROR creating UserEdits backup: " & Err.Description
    CreateUserEditsBackup = False
End Function

'===============================================================================
' USER EDITS RESTORE: Restores UserEdits data from a backup
'===============================================================================
Public Function RestoreUserEditsFromBackup(Optional backupName As String = "") As Boolean
    On Error GoTo ErrorHandler

    Dim wsEdits As Worksheet
    Dim wsBackup As Worksheet

    ' Find most recent backup if name not provided
    If backupName = "" Then
        Dim i As Integer
        For i = 1 To ThisWorkbook.Sheets.Count
            If InStr(1, ThisWorkbook.Sheets(i).Name, USEREDITS_SHEET_NAME & "_Backup_") > 0 Then ' Use Constant
                backupName = ThisWorkbook.Sheets(i).Name
                ' If multiple backups exist, we'll get the last one alphabetically
            End If
        Next i
    End If

    If backupName = "" Then
        RestoreUserEditsFromBackup = False
        Exit Function
    End If

    ' Get backup sheet
    On Error Resume Next
    Set wsBackup = ThisWorkbook.Sheets(backupName)
    On Error GoTo ErrorHandler

    If wsBackup Is Nothing Then
        RestoreUserEditsFromBackup = False
        Exit Function
    End If

    ' Ensure UserEdits sheet exists
    SetupUserEditsSheet
    Set wsEdits = ThisWorkbook.Sheets(USEREDITS_SHEET_NAME) ' Use Constant

    ' Clear current data and restore from backup
    wsEdits.Cells.Clear
    wsBackup.UsedRange.Copy wsEdits.Range("A1")

    ' Log the restoration
    LogUserEditsOperation "Restored UserEdits from backup: " & backupName

    RestoreUserEditsFromBackup = True
    Exit Function

ErrorHandler:
    Debug.Print "Error restoring from backup: " & Err.Description
    LogUserEditsOperation "ERROR restoring from backup: " & Err.Description
    RestoreUserEditsFromBackup = False
End Function

'===============================================================================
' MAIN SUB: Creates or refreshes the SQRCT Dashboard (Standardized Version)
' Optimized Restore Edits section using Array Method
' Includes logic to create/update a Text-Only version of the dashboard
' Adjusted for removed Email column
'===============================================================================
Public Sub RefreshDashboard(Optional PreserveUserEdits As Boolean = False)
    Dim ws As Worksheet, wsEdits As Worksheet
    Dim lastRow As Long, lastRowEdits As Long
    Dim docNum As String
    Dim i As Long, j As Long ' Loop counters
    Dim backupCreated As Boolean
    Dim t_start As Single, t_save As Single, t_populate As Single, t_load As Single, t_restore As Single, t_format As Single, t_textOnly As Single ' Timing variables
    Dim userEditsDict As Object ' Dictionary for UserEdits lookup (DocNum -> Sheet Row Number)
    Dim editSheetRow As Long ' Stores SHEET row number from dictionary
    Dim editArrayRow As Long ' Stores corresponding 1-based ARRAY row index
    
    ' Array variables for optimization
    Dim dashboardDocNumArray As Variant
    Dim userEditsDataArray As Variant
    Dim outputEditsArray As Variant
    Dim numDashboardRows As Long
    
    ' Variables for Text-Only sheet
    Dim wsValues As Worksheet
    Dim srcRange As Range
    Const TEXT_ONLY_SHEET_NAME As String = "SQRCT Dashboard (Text-Only)"
    Dim currentSheet As Worksheet ' To remember active sheet

    ' Create error recovery backup before any operations
    backupCreated = CreateUserEditsBackup("RefreshDashboard_" & Format(Now, "yyyymmdd_hhmmss"))
    LogUserEditsOperation "Starting dashboard refresh. PreserveUserEdits=" & PreserveUserEdits & ", Backup created: " & backupCreated

    t_start = Timer ' Start total timer
    Debug.Print "Start Refresh: " & t_start

    On Error GoTo ErrorHandler

    Application.ScreenUpdating = False
    Application.EnableEvents = False

    ' 1. Ensure UserEdits sheet exists with standardized structure
    SetupUserEditsSheet
    Set wsEdits = ThisWorkbook.Sheets(USEREDITS_SHEET_NAME) ' Use Constant

    ' 2. Save any current user edits from the dashboard to UserEdits
    '    ONLY if we're not prioritizing manually edited UserEdits
    If Not PreserveUserEdits Then
        t_save = Timer
        SaveUserEditsFromDashboard ' This now includes a dictionary load
        Debug.Print "SaveUserEdits Time: " & Timer - t_save
    End If

    ' 3. Locate/create "SQRCT Dashboard" using name constant
    Set ws = GetOrCreateDashboardSheet(DASHBOARD_SHEET_NAME) ' Use Constant and helper function

    ' *** ADDED: Ensure sheet is unprotected before any modifications ***
    On Error Resume Next ' Ignore error if already unprotected
    ws.Unprotect
    On Error GoTo ErrorHandler ' Restore error handling

    ' 4. Clean up any duplicate headers/layout issues
    CleanupDashboardLayout ws

    ' 5. Clear old data from dashboard & rebuild layout (row 3 header, etc.)
    InitializeDashboardLayout ws

    ' 6. Populate columns A-J with data from MasterQuotes_Final
    If IsMasterQuotesFinalPresent Then
        t_populate = Timer
        PopulateMasterQuotesData ws ' Uses MASTER_QUOTES_FINAL_SOURCE constant internally
        Debug.Print "PopulateFormulas Time: " & Timer - t_populate
    Else
        MsgBox "Warning: " & MASTER_QUOTES_FINAL_SOURCE & " not found. Dashboard created but no data pulled." & vbCrLf & _
               "Please ensure the data source exists.", vbInformation, "Data Source Not Found"
        GoTo Cleanup
    End If

    ' 7. Determine how many rows of data on dashboard
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    
    ' Exit restore section if no data on dashboard
    If lastRow < 4 Then
        Debug.Print "No data rows found on dashboard (lastRow=" & lastRow & "). Skipping restore."
        GoTo SkipRestore ' Skip to formatting/protection
    End If

    ' 8. Sort by First Date Pulled (F) ascending, then Document Amount (D) descending
    SortDashboardData ws, lastRow

    ' 9. AutoFit columns & fix column widths (Adjusted for removed column M)
    With ws
        .Columns("A:J").AutoFit   ' Protected columns
        .Columns(DB_COL_PHASE & ":" & DB_COL_COMMENTS).AutoFit   ' User columns (K:M)
        .Columns("C").ColumnWidth = 25  ' Customer Name
        .Columns(DB_COL_COMMENTS).ColumnWidth = 40  ' Widen User Comments (M)
    End With

    ' --- OPTIMIZED Step 10: Restore user data from UserEdits to Dashboard using Arrays ---
    t_load = Timer
    
    ' Get last row on UserEdits sheet
    lastRowEdits = wsEdits.Cells(wsEdits.Rows.Count, UE_COL_DOCNUM).End(xlUp).Row
    
    ' Read Dashboard DocNums into array
    On Error Resume Next ' Handle potential error if lastRow < 4
    dashboardDocNumArray = ws.Range("A4:A" & lastRow).Value
    If Err.Number <> 0 Then
        Debug.Print "Error reading dashboard DocNums. Skipping restore."
        Err.Clear
        GoTo SkipRestore
    End If
    On Error GoTo ErrorHandler ' Restore error handling
    
    ' Read UserEdits data into array (Columns A to F - Email D removed)
    If lastRowEdits > 1 Then
        Dim userEditsRange As Range
        ' Adjusted range to A:F (UE_COL_TIMESTAMP is now F)
        Set userEditsRange = wsEdits.Range(UE_COL_DOCNUM & "2:" & UE_COL_TIMESTAMP & lastRowEdits)
        
        If userEditsRange.Rows.Count = 1 Then ' Handle single data row case
            Dim singleRowData(1 To 1, 1 To 6) As Variant ' Create a 2D array (6 columns A-F)
            Dim cellIdx As Long
            For cellIdx = 1 To 6 ' Loop through 6 columns
                singleRowData(1, cellIdx) = userEditsRange.Cells(1, cellIdx).Value
            Next cellIdx
            userEditsDataArray = singleRowData
        Else ' Multiple rows
            userEditsDataArray = userEditsRange.Value
        End If
    Else
        Debug.Print "UserEdits sheet has no data rows (lastRowEdits=" & lastRowEdits & ")."
    End If

    ' Load Dictionary (DocNum -> Sheet Row Number)
    Set userEditsDict = LoadUserEditsToDictionary(wsEdits)
    
    Debug.Print "Load Arrays & Dictionary Time: " & Timer - t_load ' Combined time for array reads + dict load

    ' Initialize Output Array (for Dashboard columns K-M) - Reduced to 3 columns
    numDashboardRows = UBound(dashboardDocNumArray, 1)
    ReDim outputEditsArray(1 To numDashboardRows, 1 To 3) ' 3 columns: Phase, LastContact, Comments
    
    ' Pre-fill output array with blanks (vbNullString)
    For i = 1 To numDashboardRows
        For j = 1 To 3 ' Loop through 3 columns
            outputEditsArray(i, j) = vbNullString
        Next j
    Next i

    ' Process arrays to populate outputEditsArray
    t_restore = Timer
    If Not IsEmpty(userEditsDataArray) And lastRowEdits > 1 Then ' Check if UserEdits array has data
        Dim userEditsUBound As Long
        userEditsUBound = UBound(userEditsDataArray, 1)
        
        For i = 1 To numDashboardRows ' Loop through DASHBOARD rows (via array)
            docNum = Trim(CStr(dashboardDocNumArray(i, 1)))
            
            If docNum <> "" Then
                If userEditsDict.Exists(docNum) Then
                    editSheetRow = userEditsDict(docNum) ' Get SHEET row number
                    editArrayRow = editSheetRow - 1      ' Calculate corresponding ARRAY row index (SheetRow 2 = ArrayRow 1)
                    
                    ' Check if calculated array index is valid for the userEditsDataArray
                    If editArrayRow >= 1 And editArrayRow <= userEditsUBound Then
                        ' Copy data from UserEdits array to Output array
                        ' Indices adjusted for removed Email column:
                        ' Phase=2, LastContact=3, Comments=4 in userEditsDataArray
                        On Error Resume Next ' Handle potential type mismatches or errors during copy
                        outputEditsArray(i, 1) = userEditsDataArray(editArrayRow, 2) ' Phase
                        outputEditsArray(i, 2) = userEditsDataArray(editArrayRow, 3) ' LastContact
                        outputEditsArray(i, 3) = userEditsDataArray(editArrayRow, 4) ' Comments (was index 5)
                        If Err.Number <> 0 Then
                            Debug.Print "Error copying data for DocNum '" & docNum & "' at Dashboard Array Row " & i & ", UserEdits Array Row " & editArrayRow & ". Error: " & Err.Description
                            Err.Clear
                        End If
                        On Error GoTo ErrorHandler ' Restore error handling
                    Else
                         Debug.Print "Warning: DocNum '" & docNum & "' found in dictionary (Sheet Row " & editSheetRow & ") but calculated Array Row " & editArrayRow & " is out of bounds for userEditsDataArray (UBound=" & userEditsUBound & ")."
                    End If
                End If
            End If
        Next i
    Else
        Debug.Print "Skipping restore loop as userEditsDataArray is empty."
    End If
    Debug.Print "Restore Edits (Array Processing) Time: " & Timer - t_restore

    ' Write the output array back to the dashboard in one go (Adjusted range K:M)
    ws.Range(DB_COL_PHASE & "4").Resize(numDashboardRows, 3).Value = outputEditsArray
    
    ' Clean up arrays and dictionary
    Set userEditsDict = Nothing
    If IsArray(dashboardDocNumArray) Then Erase dashboardDocNumArray
    If IsArray(userEditsDataArray) Then Erase userEditsDataArray
    If IsArray(outputEditsArray) Then Erase outputEditsArray
    ' --- End OPTIMIZED Step 10 ---

SkipRestore: ' Label to jump to if restore is skipped

    ' 11. Freeze header rows
    FreezeDashboard ws

    ' 12. Apply color conditional formatting (optional) & 13. Protect columns
    t_format = Timer
    ApplyColorFormatting ws
    ProtectUserColumns ws ' Protection is now applied at the end of this sub
    Debug.Print "Format/Protect Time: " & Timer - t_format

    ' 14. Update the timestamp with improved styling
    With ws.Range("G2:I2")
        .Merge
        .Value = "Last Refreshed: " & Format$(Now(), "mm/dd/yyyy h:mm") & " MST"
        .HorizontalAlignment = xlCenter
        .Font.Size = 9
        .Font.Name = "Segoe UI"
        .Font.Color = RGB(80, 80, 80)
    End With

    ' 15. Re-create buttons with modern styling
    On Error Resume Next
    Dim shp As Shape
    For Each shp In ws.Shapes
        If shp.TopLeftCell.Row = 2 Then
            shp.Delete
        End If
    Next shp
    On Error GoTo 0

    ' Create buttons with improved styling
    ModernButton ws, "C2", "Standard Refresh", "Button_RefreshDashboard_SaveAndRestoreEdits"
    ModernButton ws, "E2", "Preserve UserEdits", "Button_RefreshDashboard_PreserveUserEdits"

    ' Build message based on which mode was used
    Dim msgText As String
    If PreserveUserEdits Then
        msgText = DASHBOARD_SHEET_NAME & " refreshed successfully!" & vbCrLf & _
                  USEREDITS_SHEET_NAME & " were preserved and applied to the dashboard." & vbCrLf & _
                  "No changes from the dashboard were saved to " & USEREDITS_SHEET_NAME & "."
    Else
        msgText = DASHBOARD_SHEET_NAME & " refreshed successfully!" & vbCrLf & _
                  "Dashboard edits were saved to " & USEREDITS_SHEET_NAME & " before refresh." & vbCrLf & _
                  USEREDITS_SHEET_NAME & " were then restored to the dashboard."
    End If

    MsgBox msgText, vbInformation, "Dashboard Refresh Complete"

    ' Log successful completion
    LogUserEditsOperation "Dashboard refresh completed successfully. Mode: " & IIf(PreserveUserEdits, "PreserveUserEdits", "StandardRefresh")
    Debug.Print "Total RefreshDashboard VBA Time: " & Timer - t_start

    ' Clean up old backups if refresh was successful
    If backupCreated Then
        On Error Resume Next
        Application.DisplayAlerts = False
        Dim oldSheets As New Collection
        Dim sh As Worksheet
        For Each sh In ThisWorkbook.Sheets
            If InStr(1, sh.Name, USEREDITS_SHEET_NAME & "_Backup_") > 0 And sh.Name <> USEREDITS_SHEET_NAME & "_Backup_" & Format(Now, "yyyymmdd") Then ' Use Constant
                oldSheets.Add sh
            End If
        Next sh

        For i = 1 To oldSheets.Count
            oldSheets(i).Delete
        Next i
        Application.DisplayAlerts = True
        On Error GoTo ErrorHandler ' Restore error handling
    End If

    ' --- Action Item 3: Create/Update Text-Only Dashboard ---
    t_textOnly = Timer ' Start Text-Only timer
    Set currentSheet = ActiveSheet ' Remember the active sheet

    On Error Resume Next ' Check if sheet exists
    Set wsValues = ThisWorkbook.Sheets(TEXT_ONLY_SHEET_NAME)
    On Error GoTo ErrorHandler ' Restore error handling

    If wsValues Is Nothing Then ' Create sheet if it doesn't exist
        Set wsValues = ThisWorkbook.Sheets.Add(After:=ws) ' Add after the main dashboard
        wsValues.Name = TEXT_ONLY_SHEET_NAME
        LogUserEditsOperation "Created sheet: " & TEXT_ONLY_SHEET_NAME
    Else ' Clear existing sheet if it exists
        wsValues.Cells.Clear
        LogUserEditsOperation "Cleared existing sheet: " & TEXT_ONLY_SHEET_NAME
    End If
    
    wsValues.Visible = xlSheetVisible ' Ensure it's visible

    ' Copy data (Values and Number Formats) from main dashboard (Adjusted range A:M)
    If lastRow >= 1 Then ' Ensure there's at least header data to copy
        Set srcRange = ws.Range("A1:" & DB_COL_COMMENTS & lastRow) ' Include headers up to new Comments col M
        srcRange.Copy
        wsValues.Range("A1").PasteSpecial Paste:=xlPasteValuesAndNumberFormats
        Application.CutCopyMode = False
        LogUserEditsOperation "Pasted values & number formats to " & TEXT_ONLY_SHEET_NAME
    End If

    ' Re-apply conditional formatting colors
    If lastRow >= 4 Then ' Only apply if there's data beyond headers
       ApplyColorFormatting wsValues ' Apply to the new sheet
       LogUserEditsOperation "Applied conditional formatting to " & TEXT_ONLY_SHEET_NAME
    End If
    
    ' Final formatting for Text-Only sheet (Adjusted range A:M)
    wsValues.Columns("A:" & DB_COL_COMMENTS).AutoFit
    
    ' Ensure sheet is unprotected
    On Error Resume Next ' In case it's already unprotected
    wsValues.Unprotect
    On Error GoTo ErrorHandler
    
    ' Ensure panes are not frozen
    wsValues.Activate ' Activate to control ActiveWindow properties
    ActiveWindow.FreezePanes = False
    currentSheet.Activate ' Re-activate the original sheet
    LogUserEditsOperation "Formatted and unfroze panes on " & TEXT_ONLY_SHEET_NAME
    Debug.Print "Create Text-Only Sheet Time: " & Timer - t_textOnly
    ' --- End Action Item 3 ---

Cleanup:
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    ' Clean up arrays just in case of early exit
    Set userEditsDict = Nothing
    If IsArray(dashboardDocNumArray) Then Erase dashboardDocNumArray
    If IsArray(userEditsDataArray) Then Erase userEditsDataArray
    If IsArray(outputEditsArray) Then Erase outputEditsArray
    Set wsValues = Nothing ' Clean up sheet object
    Set srcRange = Nothing
    Set currentSheet = Nothing
    Exit Sub

ErrorHandler:
    LogUserEditsOperation "ERROR in RefreshDashboard: " & Err.Description
    MsgBox "An error occurred during dashboard refresh. Your data has been backed up to a recovery sheet." & _
           vbCrLf & "Error: " & Err.Description, vbCritical, "Dashboard Refresh Error"
    ' Attempt to restore from backup if available
    If backupCreated Then RestoreUserEditsFromBackup
    Resume Cleanup
End Sub

'===============================================================================
' CreateSQRCTDashboard - Legacy name for backward compatibility
'===============================================================================
Public Sub CreateSQRCTDashboard()
    RefreshDashboard_TwoWaySync
End Sub

'===============================================================================
' SAVEUSEREDITSFROMDASHBOARD: Captures any user edits from Dashboard -> UserEdits
' Refactored to use Dictionary lookup for performance.
' Adjusted for removed Email column.
'===============================================================================
Public Sub SaveUserEditsFromDashboard()
    Dim wsSrc As Worksheet, wsEdits As Worksheet
    Dim lastRowSrc As Long, lastRowEdits As Long
    Dim i As Long, destRow As Long
    Dim docNum As String
    Dim hasUserEdits As Boolean
    Dim wasChanged As Boolean
    Dim userEditsDict As Object ' Dictionary for UserEdits lookup
    Dim editRow As Variant      ' To store row number or data from dictionary

    LogUserEditsOperation "Starting SaveUserEditsFromDashboard"

    ' Use GetOrCreateDashboardSheet to ensure wsSrc is valid
    Set wsSrc = GetOrCreateDashboardSheet(DASHBOARD_SHEET_NAME) ' Use Constant and helper function
    If wsSrc Is Nothing Then
        LogUserEditsOperation DASHBOARD_SHEET_NAME & " sheet not found" ' Use Constant
        Exit Sub
    End If

    SetupUserEditsSheet
    Set wsEdits = ThisWorkbook.Sheets(USEREDITS_SHEET_NAME) ' Use Constant

    lastRowSrc = wsSrc.Cells(wsSrc.Rows.Count, "A").End(xlUp).Row

    ' Load UserEdits into dictionary for faster lookup
    Set userEditsDict = LoadUserEditsToDictionary(wsEdits)
    lastRowEdits = wsEdits.Cells(wsEdits.Rows.Count, UE_COL_DOCNUM).End(xlUp).Row ' Get last row based on DocNum column
    If lastRowEdits < 1 Then lastRowEdits = 1

    For i = 4 To lastRowSrc
        docNum = Trim(CStr(wsSrc.Cells(i, "A").Value))  ' col A = Document Number
        If docNum <> "" And docNum <> "Document Number" Then

            ' Check if this row has any user edits (K-M columns now)
            hasUserEdits = False
            If wsSrc.Cells(i, DB_COL_PHASE).Value <> "" Or _
               wsSrc.Cells(i, DB_COL_LASTCONTACT).Value <> "" Or _
               wsSrc.Cells(i, DB_COL_COMMENTS).Value <> "" Then ' Email column M removed, Comments is now M
                hasUserEdits = True
            End If

            ' Find existing row using dictionary
            If userEditsDict.Exists(docNum) Then
                editRow = userEditsDict(docNum) ' Get existing row number
            Else
                editRow = 0 ' Flag as not found
            End If

            ' Process this document number if:
            ' 1. It has user edits in columns K-M, OR
            ' 2. It already exists in UserEdits (to potentially clear edits)
            If hasUserEdits Or editRow > 0 Then
                ' Determine destination row
                If editRow > 0 Then
                    destRow = editRow
                Else
                    ' Find the next available row dynamically
                    lastRowEdits = wsEdits.Cells(wsEdits.Rows.Count, UE_COL_DOCNUM).End(xlUp).Row + 1
                    If lastRowEdits < 2 Then lastRowEdits = 2 ' Ensure it's at least row 2
                    destRow = lastRowEdits
                    wsEdits.Cells(destRow, UE_COL_DOCNUM).Value = docNum
                    ' Add to dictionary immediately to handle duplicates within the dashboard itself
                    userEditsDict.Add docNum, destRow
                End If

                ' Track if we're making changes to determine if timestamp needs updating
                wasChanged = False

                ' Get current values from dashboard (Email removed)
                Dim dbPhase, dbLastContact, dbComments
                dbPhase = wsSrc.Cells(i, DB_COL_PHASE).Value
                dbLastContact = wsSrc.Cells(i, DB_COL_LASTCONTACT).Value
                ' dbEmail = wsSrc.Cells(i, DB_COL_EMAIL).Value ' REMOVED
                dbComments = wsSrc.Cells(i, DB_COL_COMMENTS).Value

                ' Only update UserEdits if either:
                ' 1. This is a new entry (editRow was 0 initially), or
                ' 2. The value in the dashboard is different from what's in UserEdits

                If editRow = 0 Or wsEdits.Cells(destRow, UE_COL_PHASE).Value <> dbPhase Then
                    wsEdits.Cells(destRow, UE_COL_PHASE).Value = dbPhase
                    wasChanged = True
                End If

                If editRow = 0 Or wsEdits.Cells(destRow, UE_COL_LASTCONTACT).Value <> dbLastContact Then
                    wsEdits.Cells(destRow, UE_COL_LASTCONTACT).Value = dbLastContact
                    wasChanged = True
                End If

                ' If editRow = 0 Or wsEdits.Cells(destRow, UE_COL_EMAIL).Value <> dbEmail Then ' REMOVED
                '     wsEdits.Cells(destRow, UE_COL_EMAIL).Value = dbEmail ' REMOVED
                '     wasChanged = True ' REMOVED
                ' End If ' REMOVED

                ' Comments column shifted from E to D in UserEdits
                If editRow = 0 Or wsEdits.Cells(destRow, UE_COL_COMMENTS).Value <> dbComments Then
                    wsEdits.Cells(destRow, UE_COL_COMMENTS).Value = dbComments
                    wasChanged = True
                End If

                ' Set ChangeSource to workbook identity and update timestamp only if something changed
                If wasChanged Then ' Update timestamp if any field was modified or if it's a new entry with edits
                    wsEdits.Cells(destRow, UE_COL_SOURCE).Value = GetWorkbookIdentity()  ' Use workbook identity (Source shifted E)
                    wsEdits.Cells(destRow, UE_COL_TIMESTAMP).Value = Format$(Now(), "yyyy-mm-dd hh:mm:ss")  ' Timestamp (Timestamp shifted F)
                    LogUserEditsOperation "Updated UserEdits for DocNumber " & docNum & " with attribution " & GetWorkbookIdentity()
                End If
            End If
        End If
'NextIteration: ' Removed GoTo
    Next i
    Set userEditsDict = Nothing ' Clean up dictionary

    LogUserEditsOperation "Completed SaveUserEditsFromDashboard"
    Exit Sub

ErrorHandler:
    LogUserEditsOperation "ERROR in SaveUserEditsFromDashboard: " & Err.Description
    ' Resume NextIteration ' Removed GoTo - Let error propagate or handle differently
    Set userEditsDict = Nothing ' Clean up dictionary on error
    ' Consider re-enabling events if necessary before exiting on error
    Application.EnableEvents = True
End Sub

'===============================================================================
' SETUPUSEREDITSSHEET: Creates or ensures existence of "UserEdits" with standard structure
' Adjusted for removed Email column.
'===============================================================================
Public Sub SetupUserEditsSheet()
    Dim wsEdits As Worksheet
    Dim needsBackup As Boolean
    Dim wsBackup As Worksheet

    LogUserEditsOperation "Starting SetupUserEditsSheet"

    ' First check if the sheet exists
    On Error Resume Next
    Set wsEdits = ThisWorkbook.Sheets(USEREDITS_SHEET_NAME) ' Use Constant
    On Error GoTo ErrorHandler

    ' If it exists, determine if we need to create a backup before modifying
    If Not wsEdits Is Nothing Then
        ' Check if we need to restructure (missing Timestamp column or wrong order, or has Email column D)
        If wsEdits.Cells(1, UE_COL_TIMESTAMP).Value <> "Timestamp" Or _
           wsEdits.Cells(1, UE_COL_SOURCE).Value <> "ChangeSource" Or _
           wsEdits.Cells(1, UE_COL_PHASE).Value <> "Engagement Phase" Or _
           wsEdits.Cells(1, 4).Value = "Email Contact" Then ' Check if old column D header exists
            needsBackup = True
            LogUserEditsOperation USEREDITS_SHEET_NAME & " sheet structure needs update - will create backup"
        End If

        ' Create backup if needed
        If needsBackup Then
            On Error Resume Next
            Set wsBackup = ThisWorkbook.Sheets(USEREDITS_SHEET_NAME & "_Backup") ' Use Constant
            If wsBackup Is Nothing Then
                Set wsBackup = ThisWorkbook.Sheets.Add(After:=wsEdits)
                wsBackup.Name = USEREDITS_SHEET_NAME & "_Backup" ' Use Constant
                LogUserEditsOperation "Created " & USEREDITS_SHEET_NAME & "_Backup sheet"
            End If
            On Error GoTo ErrorHandler

            ' Copy current data to backup
            wsEdits.UsedRange.Copy wsBackup.Range("A1")
            wsBackup.Visible = xlSheetHidden
        End If
    End If

    ' Now create or update the UserEdits sheet
    If wsEdits Is Nothing Then
        ' Sheet doesn't exist, create it new
        Set wsEdits = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        wsEdits.Name = USEREDITS_SHEET_NAME ' Use Constant
        LogUserEditsOperation "Created new " & USEREDITS_SHEET_NAME & " sheet" ' Use Constant

        ' Set up headers with improved styling using constants (6 columns A-F)
        With wsEdits.Range(wsEdits.Cells(1, UE_COL_DOCNUM), wsEdits.Cells(1, UE_COL_TIMESTAMP)) ' Use Constants for range A:F (Corrected Range)
            .Value = Array("DocNumber", "Engagement Phase", "Last Contact Date", _
                           "User Comments", "ChangeSource", "Timestamp") ' Email removed
            .Font.Bold = True
            .Interior.Color = RGB(16, 107, 193)  ' Match dashboard title
            .Font.Color = RGB(255, 255, 255)
        End With

        wsEdits.Visible = xlSheetHidden
    Else
        ' Sheet exists, make sure it has the right structure
        ' First preserve existing data if it's in the old format
        If needsBackup Then
            ' Migrate old data to new structure if needed
            If wsBackup Is Nothing Then Exit Sub  ' Safety check

            ' Capture old data structure
            Dim lastRowBackup As Long
            lastRowBackup = wsBackup.Cells(wsBackup.Rows.Count, "A").End(xlUp).Row

            ' Clear existing data
            wsEdits.Cells.Clear

            ' Set up new headers using constants (6 columns A-F)
            With wsEdits.Range(wsEdits.Cells(1, UE_COL_DOCNUM), wsEdits.Cells(1, UE_COL_TIMESTAMP)) ' Use Constants for range A:F
                .Value = Array("DocNumber", "Engagement Phase", "Last Contact Date", _
                               "User Comments", "ChangeSource", "Timestamp") ' Email removed
                .Font.Bold = True
                .Interior.Color = RGB(16, 107, 193)  ' Match dashboard title
                .Font.Color = RGB(255, 255, 255)
            End With

            ' Migrate data from backup to new structure
            If lastRowBackup > 1 Then
                Dim i As Long
                Dim oldColPhase As Long, oldColLastContact As Long, oldColEmail As Long, oldColComments As Long, oldColSource As Long, oldColTimestamp As Long
                Dim h As Long, headerText As String
                
                ' Find old column indices by header text (more robust than fixed indices)
                For h = 1 To wsBackup.UsedRange.Columns.Count
                    headerText = CStr(wsBackup.Cells(1, h).Value)
                    Select Case headerText
                        Case "UserStageOverride", "EngagementPhase", "Engagement Phase"
                            oldColPhase = h
                        Case "LastContactDate", "Last Contact Date"
                            oldColLastContact = h
                        Case "EmailContact", "Email Contact"
                            oldColEmail = h ' We find it but might not use it depending on target structure
                        Case "UserComments", "User Comments"
                            oldColComments = h
                        Case "ChangeSource" ' Handle potential old source column name
                            oldColSource = h
                        Case "Timestamp" ' Handle potential old timestamp column name
                            oldColTimestamp = h
                    End Select
                Next h

                For i = 2 To lastRowBackup
                    ' Only migrate if there's a document number
                    If wsBackup.Cells(i, 1).Value <> "" Then
                        ' Map old structure to new structure using constants, handling missing old columns gracefully
                        wsEdits.Cells(i, UE_COL_DOCNUM).Value = wsBackup.Cells(i, 1).Value  ' DocNumber (A)
                        If oldColPhase > 0 Then wsEdits.Cells(i, UE_COL_PHASE).Value = wsBackup.Cells(i, oldColPhase).Value Else wsEdits.Cells(i, UE_COL_PHASE).Value = "" ' Phase (B)
                        If oldColLastContact > 0 Then wsEdits.Cells(i, UE_COL_LASTCONTACT).Value = wsBackup.Cells(i, oldColLastContact).Value Else wsEdits.Cells(i, UE_COL_LASTCONTACT).Value = "" ' LastContact (C)
                        ' Skip Email column (old D)
                        If oldColComments > 0 Then wsEdits.Cells(i, UE_COL_COMMENTS).Value = wsBackup.Cells(i, oldColComments).Value Else wsEdits.Cells(i, UE_COL_COMMENTS).Value = "" ' Comments (D - new)
                        If oldColSource > 0 Then wsEdits.Cells(i, UE_COL_SOURCE).Value = wsBackup.Cells(i, oldColSource).Value Else wsEdits.Cells(i, UE_COL_SOURCE).Value = GetWorkbookIdentity() ' Source (E - new) - Default to current identity if missing
                        If oldColTimestamp > 0 Then wsEdits.Cells(i, UE_COL_TIMESTAMP).Value = wsBackup.Cells(i, oldColTimestamp).Value Else wsEdits.Cells(i, UE_COL_TIMESTAMP).Value = Format$(Now(), "yyyy-mm-dd hh:mm:ss") ' Timestamp (F - new) - Default to now if missing
                    End If
                Next i
            End If

            LogUserEditsOperation "Migrated " & USEREDITS_SHEET_NAME & " data to new structure (Email column removed)"
        Else
            ' Just ensure headers are correct using constants (6 columns A-F)
            With wsEdits.Range(wsEdits.Cells(1, UE_COL_DOCNUM), wsEdits.Cells(1, UE_COL_TIMESTAMP)) ' Use Constants for range A:F
                .Value = Array("DocNumber", "Engagement Phase", "Last Contact Date", _
                               "User Comments", "ChangeSource", "Timestamp") ' Email removed
                .Font.Bold = True
                .Interior.Color = RGB(16, 107, 193)
                .Font.Color = RGB(255, 255, 255)
            End With
        End If
    End If

    LogUserEditsOperation "Completed SetupUserEditsSheet"
    Exit Sub

ErrorHandler:
    LogUserEditsOperation "ERROR in SetupUserEditsSheet: " & Err.Description
    Resume Next ' Try to continue
End Sub

'===============================================================================
' GETORCREATEDASHBOARDSHEET: Returns or creates the SQRCT Dashboard
' Note: Using CodeName directly (e.g., Sheet2) is preferred if reliable.
' This function remains for cases where name-based lookup is needed.
'===============================================================================
Private Function GetOrCreateDashboardSheet(sheetName As String) As Worksheet
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(sheetName) ' Use Constant DASHBOARD_SHEET_NAME
    On Error GoTo 0

    If ws Is Nothing Then
        Set ws = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        ws.Name = sheetName
        ' Optionally set up row 1 & 2 (title, timestamp)
        SetupDashboard ws
    End If

    Set GetOrCreateDashboardSheet = ws
End Function

'===============================================================================
' CLEANUPDASHBOARDLAYOUT: Clean up duplicate rows while preserving the core structure
' Adjusted for removed Email column (now A:M)
'===============================================================================
Private Sub CleanupDashboardLayout(ws As Worksheet)
    Application.ScreenUpdating = False

    ' No need to Unprotect if UserInterfaceOnly:=True is used during Protect

    ' Step 1: Save data from row 4 onward (Adjusted range A:M)
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    Dim dataRange As Range
    Dim tempData As Variant

    If lastRow >= 4 Then
        ' Capture all data below row 3
        Set dataRange = ws.Range("A4:" & DB_COL_COMMENTS & lastRow) ' Use Comments constant (M)
        tempData = dataRange.Value
    End If

    ' Step 2: Find rows to preserve (rows 1-3)
    Dim hasTitle As Boolean
    hasTitle = False

    ' Check for title text in each cell of row 1 individually (Adjusted range A:M)
    Dim cell As Range
    For Each cell In ws.Range("A1:" & DB_COL_COMMENTS & "1").Cells ' Use Comments constant (M)
        If InStr(1, CStr(cell.Value), "STRATEGIC QUOTE RECOVERY", vbTextCompare) > 0 Then
            hasTitle = True
            Exit For
        End If
    Next cell

    ' Step 3: Clear the entire sheet EXCEPT rows 1-3 (Adjusted range A:M)
    ws.Range("A4:" & DB_COL_COMMENTS & ws.Rows.Count).Clear ' Use Comments constant (M)

    ' Step 4: If the title row (row 1) is missing, recreate it (Adjusted range A:M)
    If Not hasTitle Then
        With ws.Range("A1:" & DB_COL_COMMENTS & "1") ' Use Comments constant (M)
            .Merge
            .Value = "STRATEGIC QUOTE RECOVERY & CONVERSION TRACKER"
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
            .Font.Size = 18
            .Font.Bold = True
            .Interior.Color = RGB(16, 107, 193)  ' Slightly more vibrant blue
            .Font.Color = RGB(255, 255, 255)
            .RowHeight = 32  ' Slightly taller for a more spacious look
        End With
    End If

    ' Step 5: Ensure row 2 has control panel with professional styling (Adjusted range A:M)
    With ws.Range("A2:" & DB_COL_COMMENTS & "2") ' Use Comments constant (M)
        ' Light silver-gray gradient effect
        .Interior.Color = RGB(245, 245, 245)  ' Very light gray

        ' Add a subtle top and bottom border for definition
        .Borders(xlEdgeTop).LineStyle = xlContinuous
        .Borders(xlEdgeTop).Weight = xlThin
        .Borders(xlEdgeTop).Color = RGB(200, 200, 200)

        .Borders(xlEdgeBottom).LineStyle = xlContinuous
        .Borders(xlEdgeBottom).Weight = xlThin
        .Borders(xlEdgeBottom).Color = RGB(200, 200, 200)

        ' Set row height for better proportions
        .RowHeight = 28  ' Taller for better button spacing
    End With

    With ws.Range("A2")
        .Value = "CONTROL PANEL"
        .Font.Bold = True
        .Font.Size = 10
        .Font.Name = "Segoe UI"  ' More modern font if available
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .Interior.Color = RGB(70, 130, 180)  ' Steel blue - professional but distinct
        .Font.Color = RGB(255, 255, 255)
        .ColumnWidth = 16  ' Slightly wider for better proportions

        ' Add a subtle right border instead of surrounding border
        .Borders(xlEdgeRight).LineStyle = xlContinuous
        .Borders(xlEdgeRight).Weight = xlThin
        .Borders(xlEdgeRight).Color = RGB(200, 200, 200)
    End With

    ' Question mark in corner for help (Adjusted to column M)
    With ws.Range(DB_COL_COMMENTS & "2") ' Use Comments constant (M)
        .Value = "?"
        .Font.Bold = True
        .Font.Size = 14
        .HorizontalAlignment = xlCenter
        .Font.Color = RGB(70, 130, 180)  ' Matching steel blue
    End With

    ' Step 6: Ensure row 3 has column headers with improved styling (Adjusted range A:M)
    With ws.Range("A3:" & DB_COL_COMMENTS & "3") ' Use Comments constant (M)
        .Clear
        .Value = Array( _
            "Document Number", _
            "Client ID", _
            "Customer Name", _
            "Document Amount", _
            "Document Date", _
            "First Date Pulled", _
            "Salesperson ID", _
            "Entered By", _
            "Occurrence Counter", _
            "Missing Quote Alert", _
            "Engagement Phase", _
            "Last Contact Date", _
            "User Comments") ' Email Contact removed
        ' Headers correspond to columns:
        ' K - DB_COL_PHASE
        ' L - DB_COL_LASTCONTACT
        ' M - DB_COL_COMMENTS (Shifted from N)
        .Font.Bold = True
        .Interior.Color = RGB(16, 107, 193)  ' Match title row color
        .Font.Color = RGB(255, 255, 255)
        .HorizontalAlignment = xlCenter
    End With

    ' Step 7: Restore data if we had any (Adjusted range A:M)
    If Not IsEmpty(tempData) Then
        ' Check if tempData has the old number of columns (14) or new (13)
        Dim numCols As Long
        On Error Resume Next
        numCols = UBound(tempData, 2)
        On Error GoTo 0 ' Or appropriate error handler
        
        If numCols = 14 Then ' Old data, need to skip email column (M, index 13)
            Dim restoredData As Variant
            Dim r As Long, c As Long, targetCol As Long
            ReDim restoredData(1 To UBound(tempData, 1), 1 To 13)
            For r = 1 To UBound(tempData, 1)
                targetCol = 1
                For c = 1 To 14
                    If c <> 13 Then ' Skip old column M (index 13)
                        restoredData(r, targetCol) = tempData(r, c)
                        targetCol = targetCol + 1
                    End If
                Next c
            Next r
            ws.Range("A4").Resize(UBound(restoredData, 1), UBound(restoredData, 2)).Value = restoredData
        Else ' Assume new data structure (13 columns)
            ws.Range("A4").Resize(UBound(tempData, 1), UBound(tempData, 2)).Value = tempData
        End If
    End If

    Application.ScreenUpdating = True
End Sub

'===============================================================================
' INITIALIZEDASHBOARDLAYOUT: Clears rows 4+ in A-M, sets up header row in A3:M3
' Adjusted for removed Email column
'===============================================================================
Private Sub InitializeDashboardLayout(ws As Worksheet)
    ' Only clear rows 4+ to preserve header/control panel (Adjusted range A:M)
    ws.Range("A4:" & DB_COL_COMMENTS & ws.Rows.Count).Clear ' Use Comments constant (M)

    ' Delete extra columns N:Z if needed (Adjusted start column)
    On Error Resume Next
    ws.Range("N:Z").Delete
    On Error GoTo 0

    ' Ensure row 3 has correct headers with improved styling (Adjusted range A:M)
    With ws.Range("A3:" & DB_COL_COMMENTS & "3") ' Use Comments constant (M)
        .Clear
        .Value = Array( _
            "Document Number", _
            "Client ID", _
            "Customer Name", _
            "Document Amount", _
            "Document Date", _
            "First Date Pulled", _
            "Salesperson ID", _
            "Entered By", _
            "Occurrence Counter", _
            "Missing Quote Alert", _
            "Engagement Phase", _
            "Last Contact Date", _
            "User Comments") ' Email Contact removed
        ' Headers correspond to columns:
        ' K - DB_COL_PHASE
        ' L - DB_COL_LASTCONTACT
        ' M - DB_COL_COMMENTS (Shifted from N)
        .Font.Bold = True
        .Interior.Color = RGB(16, 107, 193)  ' Match title row color
        .Font.Color = RGB(255, 255, 255)
        .HorizontalAlignment = xlCenter
    End With

    ' Set initial column widths (Removed Email M, adjusted Comments M)
    With ws
        .Columns("A").ColumnWidth = 15
        .Columns("B").ColumnWidth = 15
        .Columns("C").ColumnWidth = 25
        .Columns("D").ColumnWidth = 15
        .Columns("E").ColumnWidth = 12
        .Columns("F").ColumnWidth = 15
        .Columns("G").ColumnWidth = 12
        .Columns("H").ColumnWidth = 15
        .Columns("I").ColumnWidth = 15
        .Columns("J").ColumnWidth = 20

        .Columns(DB_COL_PHASE).ColumnWidth = 20 ' K
        .Columns(DB_COL_LASTCONTACT).ColumnWidth = 15 ' L
        ' .Columns(DB_COL_EMAIL).ColumnWidth = 25 ' M - REMOVED
        .Columns(DB_COL_COMMENTS).ColumnWidth = 40 ' M (was N)
    End With
End Sub

'===============================================================================
' POPULATEMASTERQUOTESDATA: Pulls columns A-J from MasterQuotes_Final
' No changes needed here as it only populates A-J
'===============================================================================
Private Sub PopulateMasterQuotesData(ws As Worksheet)
    ' Use constant for the source name
    Dim sourceName As String
    sourceName = MASTER_QUOTES_FINAL_SOURCE

    With ws
        ' A: Document Number
        .Range("A4").Formula = _
            "=IF(ROWS($A$4:A4)<=ROWS(" & sourceName & "[Document Number])," & _
            "IFERROR(INDEX(" & sourceName & "[Document Number],ROWS($A$4:A4)),""""),"""")"

        ' B: Client ID -> from Customer Number
        .Range("B4").Formula = _
            "=IF(A4<>"""",IFERROR(INDEX(" & sourceName & "[Customer Number],ROWS($A$4:A4)),""""),"""")"

        ' C: Customer Name
        .Range("C4").Formula = _
            "=IF(A4<>"""",IFERROR(INDEX(" & sourceName & "[Customer Name],ROWS($A$4:A4)),""""),"""")"

        ' D: Document Amount
        .Range("D4").Formula = _
            "=IF(A4<>"""",IFERROR(--INDEX(" & sourceName & "[Document Amount],ROWS($A$4:A4)),""""),"""")"

        ' E: Document Date
        .Range("E4").Formula = _
            "=IF(A4<>"""",IFERROR(--INDEX(" & sourceName & "[Document Date],ROWS($A$4:A4)),""""),"""")"

        ' F: First Date Pulled
        .Range("F4").Formula = _
            "=IF(A4<>"""",IFERROR(--INDEX(" & sourceName & "[First Date Pulled],ROWS($A$4:A4)),""""),"""")"

        ' G: Salesperson ID
        .Range("G4").Formula = _
            "=IF(A4<>"""",IFERROR(INDEX(" & sourceName & "[Salesperson ID],ROWS($A$4:A4)),""""),"""")"

        ' H: Entered By (was User To Enter)
        .Range("H4").Formula = _
            "=IF(A4<>"""",IFERROR(INDEX(" & sourceName & "[User To Enter],ROWS($A$4:A4)),""""),"""")"

        ' I: Occurrence Counter (was Auto Stage)
        .Range("I4").Formula = _
            "=IF(A4<>"""",IFERROR(INDEX(" & sourceName & "[AutoStage],ROWS($A$4:A4)),""""),"""")"

        ' J: Missing Quote Alert (was Auto Note)
        .Range("J4").Formula = _
            "=IF(A4<>"""",IFERROR(INDEX(" & sourceName & "[AutoNote],ROWS($A$4:A4)),""""),"""")"

        ' Autofill down (adjust range if needed)
        .Range("A4:J4").AutoFill Destination:=.Range("A4:J1000"), Type:=xlFillDefault

        ' Format numeric/date columns
        .Range("D4:D1000").NumberFormat = "$#,##0.00"   ' Document Amount
        .Range("E4:E1000").NumberFormat = "mm/dd/yyyy" ' Document Date
        .Range("F4:F1000").NumberFormat = "mm/dd/yyyy" ' First Date Pulled
    End With
End Sub

'===============================================================================
' SORTDASHBOARDDATA: Sort by First Date Pulled (F asc), then Document Amount (D desc)
' Adjusted range to A:M
'===============================================================================
Private Sub SortDashboardData(ws As Worksheet, lastRow As Long)
    If lastRow < 5 Then Exit Sub

    With ws.Sort
        .SortFields.Clear
        .SortFields.Add Key:=ws.Range("F4:F" & lastRow), _
                        SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
        .SortFields.Add Key:=ws.Range("D4:D" & lastRow), _
                        SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:=xlSortNormal
        .SetRange ws.Range("A3:" & DB_COL_COMMENTS & lastRow) ' Use Comments constant (M)
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
End Sub

'===============================================================================
' FREEZEDASHBOARD: Freezes rows 1-3
' No changes needed
'===============================================================================
Private Sub FreezeDashboard(ws As Worksheet)
    ws.Activate

    ' First unfreeze any existing splits
    ActiveWindow.FreezePanes = False
    ActiveWindow.SplitRow = 0
    ActiveWindow.SplitColumn = 0

    ' Freeze rows 1-3
    ActiveWindow.SplitRow = 3
    ActiveWindow.SplitColumn = 0
    ActiveWindow.FreezePanes = True
End Sub

'===============================================================================
' SETUPDASHBOARD: Professional row 1 & 2 design (title & control panel)
' Adjusted ranges to A:M
'===============================================================================
Public Sub SetupDashboard(ws As Worksheet)
    ' Merge & style title in row 1 - updated range A:M
    With ws.Range("A1:" & DB_COL_COMMENTS & "1") ' Use Comments constant (M)
        .Merge
        .Value = "STRATEGIC QUOTE RECOVERY & CONVERSION TRACKER"
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .Font.Size = 18
        .Font.Bold = True
        .Interior.Color = RGB(16, 107, 193)  ' Slightly more vibrant blue
        .Font.Color = RGB(255, 255, 255)
        .RowHeight = 32  ' Slightly taller for a more spacious look
    End With

    ' Set up control panel in row 2 with modern styling (Adjusted range A:M)
    With ws.Range("A2:" & DB_COL_COMMENTS & "2") ' Use Comments constant (M)
        ' Light silver-gray gradient effect
        .Interior.Color = RGB(245, 245, 245)  ' Very light gray

        ' Add a subtle top and bottom border for definition
        .Borders(xlEdgeTop).LineStyle = xlContinuous
        .Borders(xlEdgeTop).Weight = xlThin
        .Borders(xlEdgeTop).Color = RGB(200, 200, 200)

        .Borders(xlEdgeBottom).LineStyle = xlContinuous
        .Borders(xlEdgeBottom).Weight = xlThin
        .Borders(xlEdgeBottom).Color = RGB(200, 200, 200)

        ' Set row height for better proportions
        .RowHeight = 28  ' Taller for better button spacing
    End With

    ' Add control panel label with professional styling
    With ws.Range("A2")
        .Value = "CONTROL PANEL"
        .Font.Bold = True
        .Font.Size = 10
        .Font.Name = "Segoe UI"  ' More modern font if available
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .Interior.Color = RGB(70, 130, 180)  ' Steel blue - professional but distinct
        .Font.Color = RGB(255, 255, 255)
        .ColumnWidth = 16  ' Slightly wider for better proportions

        ' Add a subtle right border instead of surrounding border
        .Borders(xlEdgeRight).LineStyle = xlContinuous
        .Borders(xlEdgeRight).Weight = xlThin
        .Borders(xlEdgeRight).Color = RGB(200, 200, 200)
    End With

    ' Question mark in corner for help (Adjusted to column M)
    With ws.Range(DB_COL_COMMENTS & "2") ' Use Comments constant (M)
        .Value = "?"
        .Font.Bold = True
        .Font.Size = 14
        .HorizontalAlignment = xlCenter
        .Font.Color = RGB(70, 130, 180)  ' Matching steel blue
    End With

    ' Last refreshed timestamp - positioned and styled more elegantly
    With ws.Range("G2:I2")
        .Merge
        .Value = "Last Refreshed: " & Format$(Now(), "mm/dd/yyyy h:mm") & " MST"
        .HorizontalAlignment = xlCenter
        .Font.Size = 9
        .Font.Name = "Segoe UI"
        .Font.Color = RGB(80, 80, 80)  ' Dark gray for subtle elegance
    End With

    ' Create buttons with improved spacing and styling
    ModernButton ws, "C2", "Standard Refresh", "Button_RefreshDashboard_SaveAndRestoreEdits"
    ModernButton ws, "E2", "Preserve UserEdits", "Button_RefreshDashboard_PreserveUserEdits"

End Sub

'===============================================================================
' MODERNBUTTON: Creates professional, modern-looking buttons with proper spacing
' No changes needed
'===============================================================================
Public Sub ModernButton(ws As Worksheet, cellRef As String, buttonText As String, macroName As String)
    Dim btn As Object
    Dim buttonTop As Double, buttonLeft As Double
    Dim buttonWidth As Double, buttonHeight As Double

    ' Get the position and size based on the cell
    buttonLeft = ws.Range(cellRef).Left
    buttonTop = ws.Range(cellRef).Top

    ' Calculate precise width and position to prevent overlap
    buttonWidth = ws.Range(cellRef & ":" & cellRef).Width * 1.6
    buttonHeight = ws.Range(cellRef).Height * 0.75

    ' Better centering with more space between buttons
    buttonTop = buttonTop + (ws.Range(cellRef).Height - buttonHeight) / 2

    ' Create the button using simple AddShape (9 = rounded rectangle)
    On Error Resume Next
    Set btn = ws.Shapes.AddShape(9, buttonLeft, buttonTop, buttonWidth, buttonHeight)

    If Err.Number <> 0 Then
        ' Try alternative approach for older Excel versions
        Err.Clear
        On Error Resume Next
        Set btn = ws.Shapes.AddShape(msoShapeRoundedRectangle, buttonLeft, buttonTop, buttonWidth, buttonHeight)
    End If
    On Error GoTo 0

    ' Exit if button creation failed
    If btn Is Nothing Then Exit Sub

    ' Style the button with modern, professional appearance
    With btn
        ' Modern gradient effect using solid color
        .Fill.ForeColor.RGB = RGB(42, 120, 180)  ' Professional blue

        ' More subtle border
        .Line.ForeColor.RGB = RGB(25, 95, 150)
        .Line.Weight = 0.75

        ' Set the text with improved font
        On Error Resume Next
        .TextFrame.Characters.Text = buttonText
        .TextFrame.Characters.Font.Color = RGB(255, 255, 255)
        .TextFrame.Characters.Font.Size = 10

        ' Try to use a modern font if available
        .TextFrame.Characters.Font.Name = "Segoe UI"
        .TextFrame.Characters.Font.Bold = True

        ' Text positioning
        On Error Resume Next
        .TextFrame.HorizontalAlignment = 2  ' Center = 2
        .TextFrame.VerticalAlignment = 3    ' Middle = 3

        ' Add a subtle shadow effect if possible
        On Error Resume Next
        .Shadow.Type = msoShadow21
        .Shadow.Transparency = 0.7

        ' Assign macro
        .OnAction = macroName
    End With
End Sub

'===============================================================================
' PROTECTUSERCOLUMNS: Lock A-J, unlock K-M
' Adjusted for removed Email column
'===============================================================================
Public Sub ProtectUserColumns(ws As Worksheet)
    On Error Resume Next ' Ignore errors if sheet is already unprotected
    ws.Unprotect ' Unprotect first
    On Error GoTo 0 ' Resume default error handling

    ws.Cells.Locked = True
    ' Use constants for columns (Unlock K:M)
    ws.Range(DB_COL_PHASE & "4:" & DB_COL_COMMENTS & ws.Rows.Count).Locked = False ' Use Comments constant (M)

    ' Re-apply protection here after setting Locked status
    ws.Protect UserInterfaceOnly:=True, DrawingObjects:=True, Contents:=True, Scenarios:=True
End Sub

'===============================================================================
' APPLYCOLORFORMATTING: For coloring columns I (Occurrence) and K (Engagement)
' Adjusted range to K:M
'===============================================================================
Public Sub ApplyColorFormatting(ws As Worksheet)
    On Error Resume Next ' Ignore errors if sheet is already unprotected
    ws.Unprotect ' Unprotect first
    On Error GoTo 0 ' Resume default error handling

    ' Clear existing rules in I4:I1000 and K4:M1000 (Adjusted end column)
    ws.Range("I4:I1000," & DB_COL_PHASE & "4:" & DB_COL_COMMENTS & "1000").FormatConditions.Delete ' Use Comments constant (M)

    Dim rngOccur As Range, rngPhase As Range
    Set rngOccur = ws.Range("I4:I1000")
    Set rngPhase = ws.Range(DB_COL_PHASE & "4:" & DB_COL_COMMENTS & "1000") ' Use Comments constant (M)

    ' Apply conditional formatting ONLY to the Engagement Phase column (K)
    ' ApplyStageFormatting rngOccur ' Removed - Applying text rules to numeric Col I caused errors
    ApplyStageFormatting ws.Range(DB_COL_PHASE & "4:" & DB_COL_PHASE & "1000") ' Apply only to Phase column K

    ' Re-protect, ensuring UserInterfaceOnly is True
    ws.Protect UserInterfaceOnly:=True, DrawingObjects:=True, Contents:=True, Scenarios:=True
End Sub

' Helper for detailed color rules implementing the evidence-based color system
Private Sub ApplyStageFormatting(rng As Range)
    Dim formulaBase As String
    formulaBase = "=EXACT(" & rng.Cells(1).Address(False, False) & ",""{PHASE}"")"

    ' Clear existing rules first to ensure a clean slate
    rng.FormatConditions.Delete

    With rng
        ' --- Follow-up Stages (Sequential Process) ---
        ' First F/U: Light blue (#D0E6F5)
        .FormatConditions.Add Type:=xlExpression, Formula1:=Replace(formulaBase, "{PHASE}", "First F/U")
        .FormatConditions(.FormatConditions.Count).Interior.Color = RGB(208, 230, 245)

        ' Second F/U: Medium blue (#92C6ED)
        .FormatConditions.Add Type:=xlExpression, Formula1:=Replace(formulaBase, "{PHASE}", "Second F/U")
        .FormatConditions(.FormatConditions.Count).Interior.Color = RGB(146, 198, 237)

        ' Third F/U: Yellow (#F5E171)
        .FormatConditions.Add Type:=xlExpression, Formula1:=Replace(formulaBase, "{PHASE}", "Third F/U")
        .FormatConditions(.FormatConditions.Count).Interior.Color = RGB(245, 225, 113)

        ' Long-Term F/U: Orange (#FF9636)
        .FormatConditions.Add Type:=xlExpression, Formula1:=Replace(formulaBase, "{PHASE}", "Long-Term F/U")
        .FormatConditions(.FormatConditions.Count).Interior.Color = RGB(255, 150, 54)

        ' --- Queue/Processing Status ---
        ' Requoting: Light lavender (#E3D7E8)
        .FormatConditions.Add Type:=xlExpression, Formula1:=Replace(formulaBase, "{PHASE}", "Requoting")
        .FormatConditions(.FormatConditions.Count).Interior.Color = RGB(227, 215, 232)

        ' Pending: Pale yellow (#FFF7D1)
        .FormatConditions.Add Type:=xlExpression, Formula1:=Replace(formulaBase, "{PHASE}", "Pending")
        .FormatConditions(.FormatConditions.Count).Interior.Color = RGB(255, 247, 209)

        ' No Response: Beige (#F5EEE0)
        .FormatConditions.Add Type:=xlExpression, Formula1:=Replace(formulaBase, "{PHASE}", "No Response")
        .FormatConditions(.FormatConditions.Count).Interior.Color = RGB(245, 238, 224)

        ' Texas (No F/U): Tan (#E6D9CC)
        .FormatConditions.Add Type:=xlExpression, Formula1:=Replace(formulaBase, "{PHASE}", "Texas (No F/U)")
        .FormatConditions(.FormatConditions.Count).Interior.Color = RGB(230, 217, 204)

        ' --- Team Member Assignments ---
        ' AF: Teal (#A2D9D2)
        .FormatConditions.Add Type:=xlExpression, Formula1:=Replace(formulaBase, "{PHASE}", "AF")
        .FormatConditions(.FormatConditions.Count).Interior.Color = RGB(162, 217, 210)

        ' RZ: Periwinkle (#8A9BD4)
        .FormatConditions.Add Type:=xlExpression, Formula1:=Replace(formulaBase, "{PHASE}", "RZ")
        .FormatConditions(.FormatConditions.Count).Interior.Color = RGB(138, 155, 212)

        ' KMH: Salmon (#F7C4AF)
        .FormatConditions.Add Type:=xlExpression, Formula1:=Replace(formulaBase, "{PHASE}", "KMH")
        .FormatConditions(.FormatConditions.Count).Interior.Color = RGB(247, 196, 175)

        ' RI: Sky blue (#BFE1F3)
        .FormatConditions.Add Type:=xlExpression, Formula1:=Replace(formulaBase, "{PHASE}", "RI")
        .FormatConditions(.FormatConditions.Count).Interior.Color = RGB(191, 225, 243)

        ' WW/OM: Deep purple (#9B7CB9)
        .FormatConditions.Add Type:=xlExpression, Formula1:=Replace(formulaBase, "{PHASE}", "WW/OM")
        .FormatConditions(.FormatConditions.Count).Interior.Color = RGB(155, 124, 185)

        ' --- Outcome Statuses ---
        ' Converted: Green (Adjusted x6) (Bold)
        .FormatConditions.Add Type:=xlExpression, Formula1:=Replace(formulaBase, "{PHASE}", "Converted")
        With .FormatConditions(.FormatConditions.Count)
            .Interior.Color = RGB(120, 235, 120) ' Final final final green adjustment
            .Font.Bold = True
        End With

        ' Declined: True red (#D12F2F)
        .FormatConditions.Add Type:=xlExpression, Formula1:=Replace(formulaBase, "{PHASE}", "Declined")
        .FormatConditions(.FormatConditions.Count).Interior.Color = RGB(209, 47, 47)

        ' Closed (Extra Order): Medium-dark red (#B82727)
        .FormatConditions.Add Type:=xlExpression, Formula1:=Replace(formulaBase, "{PHASE}", "Closed (Extra Order)")
        .FormatConditions(.FormatConditions.Count).Interior.Color = RGB(184, 39, 39)

        ' Closed: Dark red (#A61C1C)
        .FormatConditions.Add Type:=xlExpression, Formula1:=Replace(formulaBase, "{PHASE}", "Closed")
        .FormatConditions(.FormatConditions.Count).Interior.Color = RGB(166, 28, 28)

    End With
End Sub

'===============================================================================
' ISMASTERQUOTESFINALPRESENT: Checks for PQ, Table, or Named Range "MasterQuotes_Final"
'===============================================================================
Public Function IsMasterQuotesFinalPresent() As Boolean
    Dim lo As ListObject
    Dim nm As Name
    Dim queryObj As Object

    IsMasterQuotesFinalPresent = False

    On Error Resume Next

    ' 1) Power Query named MASTER_QUOTES_FINAL_SOURCE
    For Each queryObj In ActiveWorkbook.Queries
        If queryObj.Name = MASTER_QUOTES_FINAL_SOURCE Then ' Use Constant
            IsMasterQuotesFinalPresent = True
            Exit For
        End If
    Next queryObj

    ' 2) ListObject named MASTER_QUOTES_FINAL_SOURCE
    If Not IsMasterQuotesFinalPresent Then
        For Each lo In ActiveWorkbook.ListObjects
            If lo.Name = MASTER_QUOTES_FINAL_SOURCE Then ' Use Constant
                IsMasterQuotesFinalPresent = True
                Exit For
            End If
        Next lo
    End If

    ' 3) Named Range MASTER_QUOTES_FINAL_SOURCE
    If Not IsMasterQuotesFinalPresent Then
        For Each nm In ActiveWorkbook.Names
            If nm.Name = MASTER_QUOTES_FINAL_SOURCE Then ' Use Constant
                IsMasterQuotesFinalPresent = True
                Exit For
            End If
        Next nm
    End If

    On Error GoTo 0
End Function

'===============================================================================
' LOADUSEREDITSTODICTIONARY: Loads UserEdits DocNumbers and SHEET row numbers into a dictionary
' Note: Kept original logic mapping DocNum -> Sheet Row Number for simplicity in RefreshDashboard array access.
' Optimized by reading only DocNum column into array first.
'===============================================================================
Public Function LoadUserEditsToDictionary(wsEdits As Worksheet) As Object ' Changed Private to Public
    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")
    dict.CompareMode = vbTextCompare ' Case-insensitive keys

    If wsEdits Is Nothing Then
        Set LoadUserEditsToDictionary = dict
        Exit Function
    End If

    Dim lastRow As Long
    ' Determine last row based on DocNum column (UE_COL_DOCNUM = "A")
    lastRow = wsEdits.Cells(wsEdits.Rows.Count, UE_COL_DOCNUM).End(xlUp).Row

    If lastRow > 1 Then ' Check if there's data beyond the header row
        Dim i As Long
        Dim docNum As String
        Dim dataRange As Variant
        
        ' Read only the DocNumber column into an array for performance
        dataRange = wsEdits.Range(UE_COL_DOCNUM & "2:" & UE_COL_DOCNUM & lastRow).Value

        ' Handle case where only one row of data exists (returns a single value, not 2D array)
        If Not IsArray(dataRange) Then
            If lastRow = 2 Then
                docNum = Trim(CStr(dataRange))
                If docNum <> "" Then
                    If Not dict.Exists(docNum) Then
                        dict.Add docNum, 2 ' Store the actual SHEET row number
                    End If
                End If
            End If
        Else
            ' Process the 2D array
            For i = 1 To UBound(dataRange, 1) ' Loop through the array
                docNum = Trim(CStr(dataRange(i, 1)))
                If docNum <> "" Then
                    If Not dict.Exists(docNum) Then
                        ' Store the actual SHEET row number (i + 1 because array is 1-based and starts from sheet row 2)
                        dict.Add docNum, i + 1
                    Else
                        ' Handle potential duplicates within UserEdits if necessary (e.g., log warning)
                        ' For now, keep the first one found
                    End If
                End If
            Next i
        End If
    End If

    Set LoadUserEditsToDictionary = dict
End Function

</final_file_content>

Now that you have the latest state of the file, try the operation again with fewer/more precise SEARCH blocks.
(If you run into this error 3 times in a row, you may use the write_to_file tool as a fallback. Keep in mind, the write_to_file fallback is far from ideal, as this means you'll be re-writing the entire contents of the file just to make a few edits, which takes time and money. So let's bias towards using replace_in_file as effectively as possible)
</error><environment_details>
# VSCode Visible Files
../../../AppData/Local/Programs/Microsoft VS Code/src/ally_workbook/Module_Dashboard.bas
../../../AppData/Local/Programs/Microsoft VS Code/src/ally_workbook/Module_Dashboard.bas
src/ryan_workbook/Module_Dashboard.bas

# VSCode Open Tabs
src/ryan_workbook/Module_Dashboard.bas

# Current Time
4/10/2025, 2:20:30 PM (America/Phoenix, UTC-7:00)

# Context Window Usage
642,793 / 1,048.576K tokens used (61%)

# Current Mode
ACT MODE
</environment_details>

</file_content>

Now that you have the latest state of the file, try the operation again with fewer/more precise SEARCH blocks.
(If you run into this error 3 times in a row, you may use the write_to_file tool as a fallback. Keep in mind, the write_to_file fallback is far from ideal, as this means you'll be re-writing the entire contents of the file just to make a few edits, which takes time and money. So let's bias towards using replace_in_file as effectively as possible)
</error><environment_details>
# VSCode Visible Files
../../../AppData/Local/Programs/Microsoft VS Code/src/ryan_workbook/Module_Dashboard.bas
../../../AppData/Local/Programs/Microsoft VS Code/src/ryan_workbook/Module_Dashboard.bas
src/ryan_workbook/Module_Dashboard.bas

# VSCode Open Tabs
src/power_query/Query - MasterQuotes_Final.pq
src/power_query/Query - CSVQuotes.pq
src/ryan_workbook/Module_Dashboard.bas

# Current Time
4/10/2025, 7:10:59 PM (America/Phoenix, UTC-7:00)

# Context Window Usage
794,175 / 1,048.576K tokens used (76%)

# Current Mode
ACT MODE
</environment_details>

</file_content>

Now that you have the latest state of the file, try the operation again with fewer/more precise SEARCH blocks.
(If you run into this error 3 times in a row, you may use the write_to_file tool as a fallback. Keep in mind, the write_to_file fallback is far from ideal, as this means you'll be re-writing the entire contents of the file just to make a few edits, which takes time and money. So let's bias towards using replace_in_file as effectively as possible)
</error><environment_details>
# VSCode Visible Files
../../../AppData/Local/Programs/Microsoft VS Code/src/ryan_workbook/Module_Dashboard.bas
../../../AppData/Local/Programs/Microsoft VS Code/src/ryan_workbook/Module_Dashboard.bas
src/ryan_workbook/Module_Dashboard.bas

# VSCode Open Tabs
src/power_query/Query - MasterQuotes_Final.pq
src/power_query/Query - CSVQuotes.pq
src/ryan_workbook/Module_Dashboard.bas

# Current Time
4/10/2025, 7:31:59 PM (America/Phoenix, UTC-7:00)

# Context Window Usage
956,860 / 1,048.576K tokens used (91%)

# Current Mode
ACT MODE
</environment_details>
